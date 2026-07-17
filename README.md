# Skrypton

**You know VBS — that's enough.**

Skrypton translates **VBScript into readable, compilable C#** — and ships the runtime that makes the translated code behave *exactly* like the original did, quirks and all.

[![Build](https://img.shields.io/azure-devops/build/lubennaumov0149/Skrypton/3/main)](https://dev.azure.com/lubennaumov0149/Skrypton/_build/latest?definitionId=1&branchName=main)
[![Tests](https://img.shields.io/azure-devops/tests/lubennaumov0149/Skrypton/3/main)](https://dev.azure.com/lubennaumov0149/Skrypton/_build/latest?definitionId=3&branchName=main)
[![License](https://img.shields.io/badge/license-Apache%202.0-blue)](LICENSE)

---

## The problem

VBScript is on its way out. Microsoft has deprecated it and is **removing it from Windows** — the classic ASP pages, the `.vbs` automation, and the `MSScriptControl`-hosted rules engines your business quietly depends on are all living on borrowed time.

But VBScript isn't just "old syntax." It's a minefield of semantics that no naïve find-and-replace survives:

- **Variants everywhere** — `Empty`, `Null`, and `Nothing` are three *different* kinds of nothing, each with its own rules.
- **ByRef by default** — every argument is passed by reference unless you say otherwise.
- **`On Error Resume Next`** — errors don't stop the program, they get swallowed and inspected later via `Err`.
- **Late binding & default members** — `obj(0)`, `obj.Value`, and `obj` might all be the same call.
- **Type coercion that depends on which side of `=` a literal sits on.**

Rewrite that by hand across a large codebase and you *will* change behavior. Subtly. In production.

## The solution

Skrypton does the rewrite **for** you — and, just as importantly, gives the output a **runtime support library** that reproduces VBScript's Variant semantics, error trapping, and late binding at execution time. The generated C# doesn't just *look* like your script; it *runs* like it.

```
VBScript source ──▶ Skrypton ──▶ C# source ──▶ (compile) ──▶ .NET assembly
                                    │
                                    └── references Skrypton.RuntimeSupport (the `_` compat layer)
```

Three ways in, depending on how far you want to go:

| Goal | Use |
| --- | --- |
| **See the C#** a script becomes | `DefaultTranslator.TranslateWithoutScaffolding(...)` |
| **Analyse a script's structure** without translating | `DefaultTranslator.Parse(...)` → an `ICodeBlock` tree |
| **Run VBScript at runtime** like the old COM engine | `ScriptControlSupport.ScriptControlClass` (drop-in for `MSScriptControl.ScriptControl`) |

---

## Quick start

```csharp
using System.Globalization;
using Skrypton.CSharpWriter;
using Skrypton.CSharpWriter.Lists;

var vbscript = @"
    Dim name
    name = ""World""
    WScript.Echo ""Hello, "" & name
";

string csharp = DefaultTranslator.TranslateWithoutScaffolding(
    culture:               CultureInfo.CurrentCulture,
    scriptContent:         vbscript,
    externalDependencies:  NonNullImmutableList<string>.Empty.Add("WScript"),
    externalMemberMethods: [],
    suppressions:          []);

Console.WriteLine(csharp);
```

Or go end-to-end — translate **and execute** — through the `ScriptControl` emulation, the same way legacy apps hosted VBScript:

```csharp
var control = new ScriptControlSupport.ScriptControlClass { Language = "VBScript" };
control.AddObject("Host", myHostObject, addMembers: true);
control.AddCode(vbscript);
control.ExecuteStatement("DoTheThing");   // translated → compiled with Roslyn → run
```

---

## How it works

Skrypton is a small compiler. Source flows through a classic front-end → back-end pipeline:

**Front-end — parse (`LegacyParser` + `StageTwoParser`)**

```
raw string
  → StringBreaker       split out strings, comments, date literals
  → TokenBreaker        break the rest into atoms (names, operators, braces…)
  → OperatorCombiner    fold  <  + =  into  <= ,  collapse  -- ++  runs
  → NumberRebuilder     reassemble numeric literals ( .1 → 0.1 ), resolve  .
  → CodeBlockHandler    build an ICodeBlock tree (If / For / Class / Sub / …)
  → ExpressionGenerator (on demand) precedence-correct expression trees
```

**Back-end — generate C# (`CSharpWriter`)**

`OuterScopeBlockTranslator` rearranges VBScript's flat, everything-is-global world into a structured C# program, then a fan-out of small block translators (`If`, `For`, `ForEach`, `Do`, `Select`, `Class`, `Function`, `With`, `Erase`, …) emit the code. The output has a fixed shape:

```csharp
namespace TranslatedProgram
{
    public sealed class Runner : RunnerBaseT<EnvironmentReferences, GlobalReferences>
    {
        protected override void Go(...) { /* your outer-scope statements */ }
    }
    public sealed class GlobalReferences : ...        { /* global vars + functions */ }
    public sealed class EnvironmentReferences : ...   { /* host objects: WScript, Request… */ }
    // ...your translated VBScript classes
}
```

**Runtime — behave like VBScript (`RuntimeSupport`)**

Every operation the generated code performs goes through the compat layer (referenced as `_`): `_.ADD(a, b)`, `_.CALL(...)`, `_.IF(...)`, `_.HANDLEERROR(token, ...)`. This is where the Variant rules, `On Error Resume Next`, late binding, `CreateObject`, and 100+ built-in functions (`CInt`, `Mid`, `InStr`, `IsNull`, `CreateObject`, `FormatDateTime`, …) actually live.

---

## What's there — and what's still missing

Skrypton is real and working for the translate-and-run path, but it is **not yet a finished, packaged product.** Honest status:

### ✅ Working today

- **Full VBScript front-end** — tokenizer, expression parser, and block parser covering `If` / `ElseIf` / `Else`, `For` / `For Each`, `Do` / `While` / `Loop` / `Wend`, `Select Case`, `Class`, `Function` / `Sub` / `Property` (Get/Let/Set), `Dim` / `ReDim Preserve` / `Const` / `Public` / `Private`, `Erase`, `Exit`, `On Error Resume Next` / `Goto 0`, `Randomize`, `With`, `Option Explicit`.
- **C# code generation** for all of the above, including the hard parts: ByRef argument aliasing (VBScript ByRef ↔ C# lambdas), `On Error` trapping, deterministic `Class_Terminate` → `IDisposable`, and VBScript's literal-driven comparison coercion.
- **Runtime support library** — Variant semantics (`Empty`/`Null`/`Nothing`), the full arithmetic/logical/comparison operator set, 100+ built-in functions, the VBScript error hierarchy, and late/`IDispatch` binding.
- **`ScriptControl` emulation** — `AddCode` / `AddObject` / `ExecuteStatement` translate → compile-with-Roslyn → execute, as a managed stand-in for `MSScriptControl.ScriptControl`.
- **`Parse(...)`** for structural analysis of a script without translating it.

### 🚧 In progress / partial

- **`Executable` output mode** — the full self-contained program emitter exists and is used internally by the `ScriptControl` path, but the public `TranslateExecutable(...)` entry point is still commented out; only `TranslateWithoutScaffolding(...)` is exposed.
- **`ScriptControl` surface** — the core execute path works, but several COM members (`Eval`, `Run`, `Reset`, `Modules`, `CodeObject`, `UseSafeSubset`) are `NotImplementedException` stubs.
- **`CreateObject` / COM instantiation** — the ProgID factory plumbing is in place, but live COM activation is stubbed; you register managed factories for the objects you need.
- **Namespace cleanup** — internals still carry the upstream `VBScriptTranslator.*` namespaces (see credits) alongside the `Skrypton` public API.

### ❌ Not there yet

- **No published NuGet package** and no CLI — consume it as source / a project reference for now.
- **`RuntimeSupport/Components`** is empty — a placeholder, no drop-in components yet.
- **No migration guide / cookbook** for common patterns (see `todos.txt` for the running list of edge cases being worked through).

---

## Project layout

```
src/Skrypton/
├── LegacyParser/          front-end stage 1: tokenizing + block parsing
├── StageTwoParser/        front-end stage 2: number/operator combining + expressions
├── CSharpWriter/          back-end: ICodeBlock tree → C# source
│   └── CodeTranslation/   scope tracking, block & statement translators
├── RuntimeSupport/        the `_` compat layer the generated code runs against
└── ScriptControlSupport/  managed MSScriptControl.ScriptControl replacement
```

## Building

Requires the **.NET 10 SDK** (pinned in `global.json`). The library itself targets `netstandard2.0`.

```bash
dotnet build Skrypton.slnx -c Release
dotnet test  Skrypton.slnx -c Release
```

## Artifacts

| Name | Version |
| --- | --- |
| _NuGet package_ | _not yet published_ |

---

## Credits

Skrypton builds on the excellent [**VBScriptTranslator**](https://github.com/productiverage/VBScriptTranslator) project by Dan Roberts, which pioneered the parser, translator, and runtime-compat approach used here. It is reworked and extended under the same **Apache-2.0** license (see [LICENSE](LICENSE)).

## License

[Apache License 2.0](LICENSE) — © BENLENA.
