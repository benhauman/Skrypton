# Skrypton
Seamlessly migrate legacy VBScript to C# — translated into readable source and executed with identical runtime behavior.

[![Build Status](https://img.shields.io/azure-devops/build/lubennaumov0149/Skrypton/3/main)](https://dev.azure.com/lubennaumov0149/Skrypton/_build/latest?definitionId=3&branchName=main) [![Test Status](https://img.shields.io/azure-devops/tests/lubennaumov0149/Skrypton/3/main)](https://dev.azure.com/lubennaumov0149/Skrypton/_build/latest?definitionId=3&branchName=main) [![Code coverage](https://img.shields.io/azure-devops/coverage/lubennaumov0149/Skrypton/3/main)](https://dev.azure.com/lubennaumov0149/Skrypton/_build/latest?definitionId=3&branchName=main)

## Artifacts
| Name | Version |
| - | - |
| Skrypton | [![NuGet](https://img.shields.io/nuget/v/Skrypton)](https://www.nuget.org/packages/Skrypton) |

## Background
VBScript is being retired. Microsoft has deprecated it and is removing it from Windows, which puts every classic ASP page, `.vbs` automation script, and `MSScriptControl`-hosted rules engine on a countdown. The obvious answer — "just rewrite it in C#" — is a trap, because VBScript's behavior is not the syntax on the page:

- **Variants** — `Empty`, `Null`, and `Nothing` are three distinct kinds of "no value", each with its own comparison and coercion rules.
- **ByRef by default** — every argument is passed by reference unless declared `ByVal`.
- **`On Error Resume Next`** — errors don't halt execution; they are swallowed and inspected later through `Err`.
- **Late binding and default members** — `obj`, `obj(0)`, and `obj.Value` can all resolve to the same call.
- **Literal-driven coercion** — a comparison can behave differently depending on which side the literal sits on.

A hand rewrite silently changes these semantics, and the regressions surface in production. Skrypton exists to remove that risk: it **translates VBScript to readable C#** and ships a **runtime support library** that reproduces VBScript's semantics at execution time, so the generated code doesn't just look like the original — it behaves like it.

```
VBScript source ──▶ Skrypton ──▶ C# source ──▶ (Roslyn) ──▶ .NET assembly
                                    │
                                    └── references Skrypton.RuntimeSupport (the `_` compat layer)
```

## Getting started
Skrypton targets `netstandard2.0` and is built with the **.NET 10 SDK** (pinned in `global.json`). Install the [NuGet package](https://www.nuget.org/packages/Skrypton):

```bash
dotnet add package Skrypton
```

or build from source:

```bash
git clone https://github.com/benhauman/Skrypton.git
dotnet build Skrypton.slnx -c Release
```

There are three entry points, depending on how far you want to go:

| Goal | API |
| - | - |
| Inspect the C# a script becomes | `DefaultTranslator.TranslateWithoutScaffolding(...)` |
| Analyse a script's structure without translating | `DefaultTranslator.Parse(...)` → an `ICodeBlock` tree |
| Translate **and run** VBScript like the old COM engine | `ScriptControlSupport.ScriptControlClass` |

## Translating a script
Given a small VBScript program:

```vbscript
Dim name
name = "World"
WScript.Echo "Hello, " & name
```

translate it to C# with `DefaultTranslator`. Any names that exist in the host environment (`WScript`, `Request`, `Response`, `Session`, …) are declared as external dependencies so the translator doesn't warn about them:

```csharp
using System.Globalization;
using Skrypton.CSharpWriter;
using Skrypton.CSharpWriter.Lists;

string csharp = DefaultTranslator.TranslateWithoutScaffolding(
    culture:               CultureInfo.CurrentCulture,
    scriptContent:         vbscript,
    externalDependencies:  NonNullImmutableList<string>.Empty.Add("WScript"),
    externalMemberMethods: [],
    suppressions:          []);
```

The generated program has a fixed shape. Outer-scope statements go into a `Runner`, global variables and functions into `GlobalReferences`, host objects into `EnvironmentReferences`, and each VBScript `Class` becomes a C# class:

```csharp
namespace TranslatedProgram
{
    public sealed class Runner : RunnerBaseT<EnvironmentReferences, GlobalReferences>
    {
        protected override void Go(...) { /* outer-scope statements */ }
    }
    public sealed class GlobalReferences : ...      { /* global vars + functions */ }
    public sealed class EnvironmentReferences : ... { /* WScript, Request, ... */ }
    // ...translated VBScript classes
}
```

Every operation routes through the runtime support layer (referenced as `_`): `_.ADD(a, b)`, `_.CALL(...)`, `_.IF(...)`, `_.HANDLEERROR(token, ...)`. That layer is where Variant rules, error trapping, late binding, `CreateObject`, and the built-in functions actually live — which is why the translated code reproduces VBScript behavior rather than merely resembling it.

## Compiling the generated code
The translated C# is ordinary source. Compile it with Roslyn (or `dotnet build`) against the `Skrypton` runtime assembly, which supplies `RunnerBase`, `GlobalReferencesBase`, `EnvironmentReferencesBase`, and the `IProvideVBScriptCompatFunctionalityToIndividualRequests` compat layer. The `ScriptControl` path (below) does exactly this in-process.

## Hosting
`ScriptControlSupport.ScriptControlClass` is a managed drop-in for the `MSScriptControl.ScriptControl` COM component that legacy applications used to host VBScript. Instead of the retired scripting engine, it **translates the added code to C#, compiles it with Roslyn, and executes it**:

```csharp
var control = new ScriptControlSupport.ScriptControlClass { Language = "VBScript" };
control.AddObject("Host", myHostObject, addMembers: true); // wired in as an EnvironmentReference
control.AddCode(vbscript);
control.ExecuteStatement("DoTheThing");                    // translate → compile → run
```

## Project status
Skrypton works for the translate-and-run path but is still under active construction. Honest status:

**Working today**
- Full VBScript front-end covering `If`/`ElseIf`/`Else`, `For`/`For Each`, `Do`/`While`/`Loop`/`Wend`, `Select Case`, `Class`, `Function`/`Sub`/`Property`, `Dim`/`ReDim Preserve`/`Const`/`Public`/`Private`, `Erase`, `Exit`, `On Error Resume Next`/`Goto 0`, `Randomize`, `With`, `Option Explicit`.
- C# code generation including the hard parts: ByRef aliasing (VBScript ByRef ↔ C# lambdas), `On Error` trapping, deterministic `Class_Terminate` → `IDisposable`, and literal-driven comparison coercion.
- Runtime support library: Variant semantics, the full operator set, 100+ built-in functions, the VBScript error hierarchy, and late/`IDispatch` binding.
- `ScriptControl` emulation (`AddCode` / `AddObject` / `ExecuteStatement`) and `Parse(...)` for structural analysis.

**In progress**
- The public `TranslateExecutable(...)` entry point is commented out; only `TranslateWithoutScaffolding(...)` is exposed (the executable emitter is used internally by the `ScriptControl` path).
- Several `ScriptControl` members (`Eval`, `Run`, `Reset`, `Modules`, `CodeObject`, `UseSafeSubset`) are `NotImplementedException` stubs.
- `CreateObject` COM activation is stubbed; register managed factories for the objects you need.
- Internals still carry the upstream `VBScriptTranslator.*` namespaces (see Credits).

**Not there yet**
- No CLI.
- No migration cookbook (see `todos.txt` for the running list of edge cases).

## Syntax reference
How Skrypton maps the front-end pipeline and the language it accepts.

### Pipeline
```
raw string
  → StringBreaker       split out strings, comments, date literals
  → TokenBreaker        break the rest into atoms (names, operators, braces…)
  → OperatorCombiner    fold  <  + =  into  <= ,  collapse  -- ++  runs
  → NumberRebuilder     reassemble numeric literals ( .1 → 0.1 ), resolve  .
  → CodeBlockHandler    build an ICodeBlock tree (If / For / Class / Sub / …)
  → ExpressionGenerator (on demand) precedence-correct expression trees
```

### Supported constructs
| Category | Constructs |
| - | - |
| Conditionals | `If` / `ElseIf` / `Else`, `Select Case` |
| Loops | `For…Next`, `For Each`, `Do…Loop`, `While…Wend` |
| Declarations | `Dim`, `ReDim [Preserve]`, `Const`, `Public`, `Private` |
| Procedures | `Function`, `Sub`, `Property Get/Let/Set` (with `Default`) |
| Types | `Class` (incl. `Class_Initialize` / `Class_Terminate`) |
| Error handling | `On Error Resume Next`, `On Error Goto 0`, `Err.Raise`, `Err.Clear` |
| Statements | `Set` / `Let`, `Call`, `Erase`, `Exit`, `Randomize`, `With`, `Option Explicit` |

### Built-in functions
100+ VBScript intrinsics, including type conversion (`CInt`, `CLng`, `CDbl`, `CStr`, `CDate`, `CBool`, `Int`, `Fix`), strings (`Len`, `Mid`, `Left`, `Right`, `InStr`, `Replace`, `Split`, `Join`, `Trim`, `LCase`, `UCase`), math (`Abs`, `Round`, `Sqr`, `Rnd`, `Sgn`), type tests (`IsNull`, `IsEmpty`, `IsNumeric`, `IsObject`, `IsArray`, `TypeName`, `VarType`), arrays (`Array`, `LBound`, `UBound`), date/time (`Now`, `DateAdd`, `DateDiff`, `DatePart`, `FormatDateTime`), and object/eval (`CreateObject`, `GetObject`, `Eval`, `Execute`, `GetRef`).

### Project layout
```
src/Skrypton/
├── LegacyParser/          front-end stage 1: tokenizing + block parsing
├── StageTwoParser/        front-end stage 2: number/operator combining + expressions
├── CSharpWriter/          back-end: ICodeBlock tree → C# source
│   └── CodeTranslation/   scope tracking, block & statement translators
├── RuntimeSupport/        the `_` compat layer the generated code runs against
└── ScriptControlSupport/  managed MSScriptControl.ScriptControl replacement
```

## Credits
Skrypton builds on the [**VBScriptTranslator**](https://github.com/productiverage/VBScriptTranslator) project by Dan Roberts, which pioneered the parser, translator, and runtime-compat approach used here. It is reworked and extended under the same Apache-2.0 license.

## License
[Apache License 2.0](LICENSE) — © BENLENA.
