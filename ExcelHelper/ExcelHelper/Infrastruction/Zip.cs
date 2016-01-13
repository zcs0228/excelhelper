using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Web;

namespace ExcelHelper.Infrastruction
{
    public class Zip
    {
        /// <summary>
        /// 压缩文件夹
        /// </summary>
        /// <param name="folderName"></param>
        /// <param name="compressedFileName"></param>
        /// <param name="overrideExisting"></param>
        /// <returns></returns>
        public static bool PackageFolder(string folderName, string compressedFileName, bool overrideExisting)
        {
            if (folderName.EndsWith(@"\"))
                folderName = folderName.Remove(folderName.Length - 1);
            bool result = false;
            if (!Directory.Exists(folderName))
            {
                return result;
            }

            if (!overrideExisting && File.Exists(compressedFileName))
            {
                return result;
            }
            try
            {
                using (Package package = Package.Open(compressedFileName, FileMode.Create))
                {
                    var fileList = Directory.EnumerateFiles(folderName, "*", SearchOption.AllDirectories);
                    foreach (string fileName in fileList)
                    {

                        //The path in the package is all of the subfolders after folderName
                        string pathInPackage;
                        pathInPackage = Path.GetDirectoryName(fileName).Replace(folderName, string.Empty) + "/" + Path.GetFileName(fileName);

                        Uri partUriDocument = PackUriHelper.CreatePartUri(new Uri(pathInPackage, UriKind.Relative));
                        PackagePart packagePartDocument = package.CreatePart(partUriDocument, "", CompressionOption.Maximum);
                        using (FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read))
                        {
                            fileStream.CopyTo(packagePartDocument.GetStream());
                        }
                    }
                }
                result = true;
            }
            catch (Exception e)
            {
                throw new Exception("Error zipping folder " + folderName, e);
            }

            return result;
        }

        /// <summary>
        /// 压缩文件
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="compressedFileName"></param>
        /// <param name="overrideExisting"></param>
        /// <returns></returns>
        public static bool PackageFile(string fileName, string compressedFileName, bool overrideExisting)
        {
            bool result = false;

            if (!File.Exists(fileName))
            {
                return result;
            }

            if (!overrideExisting && File.Exists(compressedFileName))
            {
                return result;
            }

            try
            {
                Uri partUriDocument = PackUriHelper.CreatePartUri(new Uri(Path.GetFileName(fileName), UriKind.Relative));

                using (Package package = Package.Open(compressedFileName, FileMode.OpenOrCreate))
                {
                    if (package.PartExists(partUriDocument))
                    {
                        package.DeletePart(partUriDocument);
                    }

                    PackagePart packagePartDocument = package.CreatePart(partUriDocument, "", CompressionOption.Maximum);
                    using (FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read))
                    {
                        fileStream.CopyTo(packagePartDocument.GetStream());
                    }
                }
                result = true;
            }
            catch (Exception e)
            {
                throw new Exception("Error zipping file " + fileName, e);
            }

            return result;
        }

        public static bool UncompressFile(string folderName, string compressedFileName, bool overrideExisting)
        {
            bool result = false;
            try
            {
                if (!File.Exists(compressedFileName))
                {
                    return result;
                }

                DirectoryInfo directoryInfo = new DirectoryInfo(folderName);
                if (!directoryInfo.Exists)
                    directoryInfo.Create();

                using (Package package = Package.Open(compressedFileName, FileMode.Open, FileAccess.Read))
                {
                    foreach (PackagePart packagePart in package.GetParts())
                    {
                        ExtractPart(packagePart, folderName, overrideExisting);
                    }
                }

                result = true;
            }
            catch (Exception e)
            {
                throw new Exception("Error unzipping file " + compressedFileName, e);
            }

            return result;
        }

        /// <summary>
        /// 解压文件
        /// </summary>
        /// <param name="packagePart"></param>
        /// <param name="targetDirectory"></param>
        /// <param name="overrideExisting"></param>
        public static void ExtractPart(PackagePart packagePart, string targetDirectory, bool overrideExisting)
        {
            string stringPart = targetDirectory + HttpUtility.UrlDecode(packagePart.Uri.ToString()).Replace('\\', '/');

            if (!Directory.Exists(Path.GetDirectoryName(stringPart)))
                Directory.CreateDirectory(Path.GetDirectoryName(stringPart));

            if (!overrideExisting && File.Exists(stringPart))
                return;
            using (FileStream fileStream = new FileStream(stringPart, FileMode.Create))
            {
                packagePart.GetStream().CopyTo(fileStream);
            }
        }

        /// <summary>
        /// 获得部分内容的流对象，用完后必须关闭流
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="uri"></param>
        /// <returns></returns>
        public static Stream GetPartStream(string fileName, string uri)
        {
            Stream result = null;
            Package package = Package.Open(fileName, FileMode.Open);
            PackagePartCollection parts = package.GetParts();           
            foreach(var part in parts)
            {
                string uriString = part.Uri.ToString();
                if (String.Compare(uriString, uri) == 0)
                {
                    result = part.GetStream(FileMode.Open, FileAccess.Read);
                }
            }
            return result;
        }
    }
}
