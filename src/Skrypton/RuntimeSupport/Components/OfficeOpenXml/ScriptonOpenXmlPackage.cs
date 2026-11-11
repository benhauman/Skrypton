using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace Skrypton.RuntimeSupport.Components
{
    /// <summary>
    /// An Office Open XML package (.docx, .xlsx) in memory: the parts are kept as they were read and the parts that are changed are replaced when
    /// the package is written. Files are read and written through the host's IHostFileSystemHostService.
    /// </summary>
    internal sealed class ScriptonOpenXmlPackage
    {
        private static readonly char[] PathSeparators = { '/' };

        private readonly Dictionary<string, byte[]> _parts;

        private ScriptonOpenXmlPackage(Dictionary<string, byte[]> parts)
        {
            _parts = parts;
        }

        public static ScriptonOpenXmlPackage Create(IEnumerable<KeyValuePair<string, string>> xmlParts)
        {
            var parts = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
            foreach (var part in xmlParts)
                parts[part.Key] = new UTF8Encoding(false).GetBytes(part.Value);
            return new ScriptonOpenXmlPackage(parts);
        }

        /// <summary>
        /// Read the package (InvalidDataException if it is not a zip package)
        /// </summary>
        public static ScriptonOpenXmlPackage Load(IHostFileSystemHostService fileSystem, string path)
        {
            byte[] content;
            using (StreamReader reader = fileSystem.OpenTextFileRead(path))
            using (var memory = new MemoryStream())
            {
                reader.BaseStream.CopyTo(memory);
                content = memory.ToArray();
            }

            var parts = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
            using (var archive = new ZipArchive(new MemoryStream(content), ZipArchiveMode.Read))
            {
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    if (entry.FullName.EndsWith("/", StringComparison.Ordinal))
                        continue;
                    using (Stream stream = entry.Open())
                    using (var memory = new MemoryStream())
                    {
                        stream.CopyTo(memory);
                        parts[entry.FullName] = memory.ToArray();
                    }
                }
            }
            return new ScriptonOpenXmlPackage(parts);
        }

        public void Save(IHostFileSystemHostService fileSystem, string path)
        {
            byte[] content;
            using (var memory = new MemoryStream())
            {
                using (var archive = new ZipArchive(memory, ZipArchiveMode.Create, leaveOpen: true))
                {
                    // [Content_Types].xml first, as Office writes it
                    foreach (var part in _parts.OrderBy(p => p.Key.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase) ? 0 : 1))
                    {
                        ZipArchiveEntry entry = archive.CreateEntry(part.Key, CompressionLevel.Optimal);
                        using (Stream stream = entry.Open())
                            stream.Write(part.Value, 0, part.Value.Length);
                    }
                }
                content = memory.ToArray();
            }
            using (Stream stream = fileSystem.OpenTextFileWrite(path, createIfNotExists: true, overwriteIfExists: true, append: false))
                stream.Write(content, 0, content.Length);
        }

        public ScriptonOpenXmlPackage Clone() => new ScriptonOpenXmlPackage(new Dictionary<string, byte[]>(_parts, StringComparer.OrdinalIgnoreCase));

        public bool HasPart(string name) => _parts.ContainsKey(name);

        public XDocument GetXml(string name)
        {
            if (!_parts.TryGetValue(name, out byte[]? content))
                throw new InvalidDataException($"The package has no part '{name}'");
            using (var stream = new MemoryStream(content))
                return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
        }

        public void SetXml(string name, XDocument document)
        {
            using (var memory = new MemoryStream())
            {
                using (var writer = System.Xml.XmlWriter.Create(memory, new System.Xml.XmlWriterSettings { Encoding = new UTF8Encoding(false) }))
                    document.Save(writer);
                _parts[name] = memory.ToArray();
            }
        }

        /// <summary>
        /// The target of the relationship of the given type from the part (eg. the main document part from the package relationships), relative to
        /// the package root
        /// </summary>
        public string? GetRelationshipTarget(string sourcePart, string relationshipTypeSuffix)
        {
            string relsPart = GetRelationshipsPart(sourcePart);
            if (!HasPart(relsPart))
                return null;
            XNamespace rel = "http://schemas.openxmlformats.org/package/2006/relationships";
            XElement? relationship = GetXml(relsPart).Root?.Elements(rel + "Relationship")
                .FirstOrDefault(r => ((string?)r.Attribute("Type") ?? "").EndsWith(relationshipTypeSuffix, StringComparison.Ordinal));
            string? target = (string?)relationship?.Attribute("Target");
            return (target == null) ? null : ResolvePartName(sourcePart, target);
        }

        public IEnumerable<KeyValuePair<string, string>> GetRelationships(string sourcePart)
        {
            string relsPart = GetRelationshipsPart(sourcePart);
            if (!HasPart(relsPart))
                yield break;
            XNamespace rel = "http://schemas.openxmlformats.org/package/2006/relationships";
            foreach (XElement relationship in GetXml(relsPart).Root!.Elements(rel + "Relationship"))
                yield return new KeyValuePair<string, string>((string?)relationship.Attribute("Id") ?? "", ResolvePartName(sourcePart, (string?)relationship.Attribute("Target") ?? ""));
        }

        private static string GetRelationshipsPart(string sourcePart)
        {
            if (sourcePart.Length == 0)
                return "_rels/.rels";
            int slash = sourcePart.LastIndexOf('/');
            return sourcePart.Substring(0, slash + 1) + "_rels/" + sourcePart.Substring(slash + 1) + ".rels";
        }

        private static string ResolvePartName(string sourcePart, string target)
        {
            if (target.StartsWith("/", StringComparison.Ordinal))
                return target.Substring(1);
            int slash = sourcePart.LastIndexOf('/');
            var segments = new List<string>((slash < 0 ? "" : sourcePart.Substring(0, slash)).Split(PathSeparators, StringSplitOptions.RemoveEmptyEntries));
            foreach (string segment in target.Split('/'))
            {
                if (segment == "..")
                {
                    if (segments.Count > 0)
                        segments.RemoveAt(segments.Count - 1);
                }
                else if ((segment.Length > 0) && (segment != "."))
                {
                    segments.Add(segment);
                }
            }
            return string.Join("/", segments);
        }
    }
}
