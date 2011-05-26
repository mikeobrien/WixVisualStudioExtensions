using System;

namespace WixVisualStudioExtensions
{
    public static class WixContent
    {
        public static string GenerateNewGuid()
        {
            return Guid.NewGuid().ToString("D").ToUpper();
        }

        public static string GenerateFileId()
        {
            return GenerateId(10);
        }

        public static string GenerateShortFileName()
        {
            return GenerateRandomShortFilename();
        }

        public static string GenerateComponentTag()
        {
            const string componentFormat = "\r\n<Component Id=\"\" Guid=\"{0}\">\r\n\r\n</Component>\r\n";
            return string.Format(componentFormat, Guid.NewGuid().ToString("D"));
        }

        public static string GenerateDirectoryTag()
        {
            const string directoryFormat = "\r\n <Directory Id=\"{0}\" Name=\"{1}\" LongName=\"\">\r\n\r\n</Directory>\r\n";
            return string.Format(directoryFormat, GenerateId(10), GenerateRandomShortFilename());
        }

        public static string GenerateFileTag(bool includeShortcut)
        {
            const string fileElementFormat = "\r\n<File Id=\"{0}\" Name=\"{1}\" LongName=\"\" DiskId=\"1\" Source=\"\" Vital=\"yes\"{2}";
            const string fileElementCloseFormat = " />\r\n";
            const string shortcutFormat = ">\r\n<Shortcut Id=\"{0}\" Directory=\"ProgramMenuDir|DesktopFolder\" Name=\"{1}\" LongName=\"\" WorkingDirectory=\"INSTALLDIR\" />\r\n</File>\r\n";

            return string.Format(fileElementFormat,
                                 GenerateId(10),
                                 GenerateRandomShortFilename(),
                                 !includeShortcut ? fileElementCloseFormat : string.Format(shortcutFormat, GenerateId(10), GenerateRandomShortFilename()));
        }

        private static string GenerateRandomShortFilename()
        {
            return string.Format("{0}.{1}", GenerateId(8), GenerateId(3));
        }

        private static string GenerateId(int length)
        {
            var random = new Random(DateTime.Now.Millisecond);
            var randomId = ((char)random.Next(97, 102)).ToString();
            var seed = Guid.NewGuid().ToString("N");

            for (var index = 0; index < length - 1; index++)
            {
                randomId += seed.Substring(random.Next(0, 31), 1);
            }
            return randomId;
        }
    }
}