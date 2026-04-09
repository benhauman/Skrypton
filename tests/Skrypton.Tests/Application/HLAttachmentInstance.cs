using System.Runtime.InteropServices;

namespace Skrypton.Tests.Application;

[ComVisible(true)]
public sealed class HLAttachmentInstance // see IHlAttachment
{
    private string internalName;
    private byte[] fileBytes;

    public HLAttachmentInstance(string name, string dataType, byte[] fileBytes)
    {
        internalName = name;
        this.fileBytes = fileBytes;
    }
    public object GetName()
    {
        return internalName;
    }
    public object GetData()
    {
        if (fileBytes == null)
        {
            //fileBytes = _helplineObjectPersistence.LoadAttachment(defId, id, serviceUnitAttachment);
            //Size = fileBytes.Length;
        }

        return fileBytes;
    }
    public int GetSize()
    {
        return fileBytes.Length;
    }
    public string GetURL()
    {
        return fileBytes.Length != 0 ? "" : internalName;
    }
    /*
[DispId(1)]
           [MethodImpl(MethodImplOptions.InternalCall)]
           [return: MarshalAs(UnmanagedType.Struct)]
           object GetID();
           [DispId(2)]
           [MethodImpl(MethodImplOptions.InternalCall)]
           [return: MarshalAs(UnmanagedType.Struct)]
           object GetName();
           [DispId(3)]
           [MethodImpl(MethodImplOptions.InternalCall)]
           [return: MarshalAs(UnmanagedType.Struct)]
           object GetSize();
           [DispId(4)]
           [MethodImpl(MethodImplOptions.InternalCall)]
           [return: MarshalAs(UnmanagedType.Struct)]
           object GetURL();
           [DispId(5)]
           [MethodImpl(MethodImplOptions.InternalCall)]
           [return: MarshalAs(UnmanagedType.Struct)]
           object GetLastModified();
           [DispId(6)]
           [MethodImpl(MethodImplOptions.InternalCall)]
           [return: MarshalAs(UnmanagedType.Struct)]
           object GetData();
           [DispId(7)]
           [MethodImpl(MethodImplOptions.InternalCall)]
           void SetID([In] int val);
           [DispId(8)]
           [MethodImpl(MethodImplOptions.InternalCall)]
           void SetName([MarshalAs(UnmanagedType.Struct)] [In] object val);
           [DispId(9)]
           [MethodImpl(MethodImplOptions.InternalCall)]
           void SetSize([In] int val);
           [DispId(10)]
           [MethodImpl(MethodImplOptions.InternalCall)]
           void SetURL([MarshalAs(UnmanagedType.Struct)] [In] object val);
           [DispId(11)]
           [MethodImpl(MethodImplOptions.InternalCall)]
           void SetLastModified([In] int val);
           [DispId(12)]
           [MethodImpl(MethodImplOptions.InternalCall)]
           void SetData([MarshalAs(UnmanagedType.Struct)] [In] object val, [In] int size);
           [DispId(13)]
           [MethodImpl(MethodImplOptions.InternalCall)]
           void SetFile([MarshalAs(UnmanagedType.Struct)] [In] object val);
           [DispId(14)]
           [MethodImpl(MethodImplOptions.InternalCall)]
           [return: MarshalAs(UnmanagedType.Struct)]
           object GetFile();
           [DispId(15)]
           [MethodImpl(MethodImplOptions.InternalCall)]
           void SetDataWithCheck([MarshalAs(UnmanagedType.Struct)] [In] object val, [In] int size);


           string ContentId { get; set; }
           bool IsEmbedded { get; set; }
           string Name { get; set; }
           int Id { get; set; }
           long Size { get; set; }
           DateTime UploadTime { get; set; }
           string Extension { get; set; }

           void InitializeDependencies(IHelplineObjectPersistence helplineObjectPersistence);
         */


    /*
        public object GetID()
               {
                   return id;
               }

               public object GetSize()
               {
                   return (int)Size;
               }

               public object GetURL()
               {
                   if (Size != 0)
                       return "";
                   else
                       return internalName;
               }

               public object GetLastModified()
               {
                   return lastModified.ToOADate();
               }


               public void SetID(int val)
               {
                   this.id = val;
               }

               public void SetName(object val)
               {
                   this.internalName = (string)val;
               }

               public void SetSize(int val)
               {
                   this.Size = val;
               }

               public void SetURL(object val)
               {
                   internalName = (string)val;
               }

               public void SetLastModified(int val)
               {
                   lastModified = new System.DateTime(1970, 1, 1).AddSeconds(val);
               }

               public void SetData(object val, int size)
               {
                   SetDataInternal(val, size, false);
               }

               public void SetDataWithCheck(object val, int size)
               {
                   SetDataInternal(val, size, true);
               }

               private void SetDataInternal(object val, int size, bool check)
               {
                   fileBytes = (byte[])val;

                   if (check)
                   {
                       if (fileBytes == null || fileBytes.Length == 0)
                       {
                           throw new ArgumentException(StringTable.CannotAddEmptyAttachment);
                       }

                       int maximumAttachmentSizeInMB = 20; // Maximum file Size 20 MB
                       if (fileBytes.Length > maximumAttachmentSizeInMB * 1024 * 1024)
                       {
                           throw new ArgumentException(string.Format(StringTable.AttachmentTooLarge, maximumAttachmentSizeInMB));
                       }
                   }

                   Size = size;
                   lastModified = DateTime.Now;
               }

               public void SetFile(object val)
               {
                   SetName(Path.GetFileName((string)val));

                   var bytes = ReadAllBytes((string)val);
                   SetDataWithCheck(bytes, bytes.Length);
               }

               // Well we need to open the FileStream while reading with a share mode of "ReadWrite" instead of "Read",
               // so we can not use the default .Net implementation of ReadAllBytes.
               private static byte[] ReadAllBytes(String path)
               {
                   byte[] bytes;
                   using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                   {
                       int index = 0;
                       long fileLength = fs.Length;
                       if (fileLength > Int32.MaxValue)
                           throw new IOException("File too long");
                       int count = (int)fileLength;
                       bytes = new byte[count];
                       while (count > 0)
                       {
                           int n = fs.Read(bytes, index, count);
                           if (n == 0)
                               throw new InvalidOperationException("End of file reached before expected");
                           index += n;
                           count -= n;
                       }
                   }
                   return bytes;
               }

               /// <summary>
               /// Speichert das Attachment in eine temporäre Datei und liefert den Dateinamen zurück.
               /// Die temporäre Datei wird im Dispose oder im Destruktor gelöscht.
               /// </summary>
               /// <returns>Der vollständige Pfad zur temporären Datei.</returns>
               public object GetFile()
               {
                   if (this.fileBytes == null)
                   {
                       GetData();
                   }

                   string tempDirectory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
                   Directory.CreateDirectory(tempDirectory);

                   this.tempFilePath = Path.Combine(tempDirectory, Attachment.ReplaceIllegalFileNameChars(this.Name, "_"));

                   if (tempFilePath.Length >= 260)         // If Path is too long, create new filename with generated guid filename and correct extension
                       this.tempFilePath = Path.Combine(tempDirectory, Guid.NewGuid().ToString() + Path.GetExtension(Name));

                   File.WriteAllBytes(this.tempFilePath, this.fileBytes);

                   return this.tempFilePath;
               }

               /// <summary>
               /// Replaces all illegal characters for file names
               /// </summary>
               /// <param name="filename">the file name that could contain illegal characters</param>
               /// <param name="replacement">the replacement for illegal characters</param>
               /// <returns></returns>
               public static string ReplaceIllegalFileNameChars(string filename, string replacement)
               {
                   string regexSearch = new string(Path.GetInvalidFileNameChars());
                   var regex = new System.Text.RegularExpressions.Regex(string.Format("[{0}]", System.Text.RegularExpressions.Regex.Escape(regexSearch)));
                   return regex.Replace(filename, replacement);
               }

               public void Dispose()
               {
                   DeleteTempFile();
               }

               private void DeleteTempFile()
               {
                   if (!string.IsNullOrEmpty(tempFilePath) && File.Exists(tempFilePath))
                   {
                       try
                       {
                           File.Delete(tempFilePath);

                           // Auch noch versuchen, das Verzeichnis zu löschen...
                           string tempPath = Path.GetDirectoryName(tempFilePath);
                           Directory.Delete(tempPath);

                           tempFilePath = "";
                       }
                       catch
                       {
                       }
                   }
               }

               public string ContentId
               {
                   get
                   {
                       if (internalName == null)
                           return null;
                       return AttachmentFileNameSerializer.Deserialize(internalName).ContentId;
                   }
                   set
                   {
                       internalName = AttachmentFileNameSerializer.Serialize(Name, value, IsEmbedded);
                   }
               }

               public bool IsEmbedded
               {
                   get
                   {
                       if (internalName == null)
                           return false;

                       return AttachmentFileNameSerializer.Deserialize(internalName).IsEmbedded;
                   }
                   set
                   {
                       internalName = AttachmentFileNameSerializer.Serialize(Name, ContentId, value);
                  }
               }

               public string Name
               {
                   get
                   {
                       if (internalName == null)
                           return null;

                       return AttachmentFileNameSerializer.Deserialize(internalName).FileName;
                   }
                   set
                   {
                       internalName = AttachmentFileNameSerializer.Serialize(value, ContentId, IsEmbedded);
                   }
               }

               public int Id
               {
                   get
                   {
                       return id;
                   }
                   set
                   {
                       id = value;
                   }
               }

               public long Size { get; set; }

               public DateTime UploadTime
               {
                   get
                   {
                       return lastModified;
                   }
                   set
                   {
                       lastModified = value;
                   }
               }

               public string Extension { get; set; }

               public override int GetHashCode()
               {
                   return this.Id.GetHashCode();
               }

               public override bool Equals(object obj)
               {
                   Attachment attachment = (Attachment)obj;

                   if (attachment.Id != this.Id)
                       return false;

                   return true;
               }

               public void InitializeDependencies(IHelplineObjectPersistence helplineObjectPersistence)
               {
                   this._helplineObjectPersistence = helplineObjectPersistence;
               }
           }

           public class AttachmentKey
           {
               public AttachmentKey()
               {
               }

               public AttachmentKey(string attributeKey, int serviceUnit) : this()
               {
                   this.AttributeKey = attributeKey;
                   this.ServiceUnitIndex = serviceUnit;
               }

               public string AttributeKey { get; private set; }

               public int ServiceUnitIndex { get; private set; }

               public override int GetHashCode()
               {
                   string hash = string.Format("{0}_{1}", AttributeKey, ServiceUnitIndex);
                   return hash.GetHashCode();
               }

               public override bool Equals(object obj)
               {
                   AttachmentKey ak = (AttachmentKey)obj;
                   if (ak.AttributeKey != AttributeKey)
                       return false;
                   if (ak.ServiceUnitIndex != ServiceUnitIndex)
                       return false;

                   return true;
               }

               public override string ToString()
               {
                   return string.Format("{0}; {1}", AttributeKey, ServiceUnitIndex);
               }
           }
         */
}