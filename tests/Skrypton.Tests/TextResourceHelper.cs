namespace Skrypton.Tests
{
    using System;
    using System.IO;
    using System.Linq;
    using System.Reflection;

    public static class TextResourceHelper
    {


        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1004:GenericMethodsShouldProvideTypeParameter")]
        public static string LoadResourceText<T>(string resourceName)
        {
            return LoadResourceText(typeof(T), resourceName);
        }

        public static string LoadResourceText(Type typeFromResourceAssembly, string resourceName)
        {
            using (TextReader textReader = LoadResourceString(typeFromResourceAssembly, resourceName))
            {
                return textReader.ReadToEnd();
            }
        }
        public static TextReader LoadResourceString(Type typeFromResourceAssembly, string resourceName)
        {
            Stream resourceStream = GetResourceStream(typeFromResourceAssembly, resourceName);
            TextReader textReader = new StreamReader(resourceStream);
            return textReader;
        }

        public static Stream GetResourceStream(Type typeResourceAssembly, string resourceName)
        {
            if (typeResourceAssembly == null)
                throw new ArgumentNullException("typeResourceAssembly");
            if (string.IsNullOrEmpty(resourceName))
                throw new ArgumentNullException("resourceName");

            Assembly resourceAssembly = typeResourceAssembly.Assembly;
            Stream resourceStream = resourceAssembly.GetManifestResourceStream(resourceName);
            if (resourceStream == null)
            {
                string[] names = resourceAssembly.GetManifestResourceNames()
                    .OrderBy(x => x).ToArray();

                if (names != null) { }

                foreach (var resAsm_name in names)
                {
                    int cmp = string.Compare(resAsm_name, resourceName, StringComparison.Ordinal);
                    if (cmp == 0)
                    {
                        break;
                    }
                    else
                    {
                        // set the next statement here if the length are equal
                        if (resAsm_name.Length == resourceName.Length)
                        {
                            for (int c_ix = 0; c_ix < resAsm_name.Length; c_ix++)
                            {
                                var resAsm_name_c = resAsm_name[c_ix];
                                var resourcName_c = resourceName[c_ix];
                                if (resAsm_name_c == resourcName_c)
                                {
                                }
                                else
                                {
                                    break; // !! this character is the problem
                                }
                            }
                        }

                    }
                }

                throw new InvalidOperationException("Resource could not be found:[" + resourceName + "].");// Available resources:" + string.Join(",", names));
            }
            else
                return resourceStream;
        }
    }
}
