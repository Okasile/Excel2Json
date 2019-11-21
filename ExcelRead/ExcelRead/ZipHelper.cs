using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ICSharpCode.SharpZipLib.Core;
using ICSharpCode.SharpZipLib.Zip;
using ICSharpCode.SharpZipLib.GZip;
using ICSharpCode.SharpZipLib.Tar;


namespace ExcelRead
{
    class ZipHelper
    {
        static ZipHelper()
        {
            ZipStrings.CodePage = Encoding.UTF8.CodePage;
        }


        public static byte[] ZipCompress(byte[] bytes, string zipEntryName, int level = 3)
        {
            MemoryStream msIn = new MemoryStream(bytes);
            MemoryStream msOut = CreateToMemoryStream(msIn, zipEntryName, level);

            byte[] result = msOut.ToArray();

            msIn.Close();
            msOut.Close();
            return result;
        }

        public static byte[] ZipDecompress(byte[] bytes)
        {
            string errorMsg;
            return ZipDecompress(bytes, out errorMsg);
        }
        public static byte[] ZipDecompress(byte[] bytes, out string errorMsg)
        {
            byte[] result = null;
            errorMsg = null;

            if (bytes != null)
            {
                MemoryStream msIn = new MemoryStream(bytes);
                MemoryStream msOut = null;
                try
                {
                    msOut = UnzipFromStream(msIn);
                    result = msOut.ToArray();
                }
                catch (Exception e)
                {
                    result = null;
                    errorMsg = e.Message;
                }
                finally
                {
                    msIn.Close();
                    if (msOut != null)
                        msOut.Close();
                }
            }
            return result;
        }




        public static byte[] GZipCompress(byte[] bytes, int level = 3)
        {
            MemoryStream msIn = new MemoryStream(bytes);
            MemoryStream msOut = gZip(msIn, level);

            byte[] result = msOut.ToArray();

            msIn.Close();
            msOut.Close();
            return result;
        }
        public static MemoryStream GZipCompress(MemoryStream inStream, int level = 3)
        {
            return gZip(inStream, level);
        }

        public static byte[] GZipDecompress(byte[] bytes)
        {
            string errorMsg;
            return GZipDecompress(bytes, out errorMsg);
        }
        public static byte[] GZipDecompress(byte[] bytes, out string errorMsg)
        {
            byte[] result = null;
            errorMsg = null;

            if (bytes != null)
            {
                MemoryStream msIn = new MemoryStream(bytes);
                MemoryStream msOut = null;
                try
                {
                    msOut = gunZip(msIn);
                    result = msOut.ToArray();
                }
                catch (Exception e)
                {
                    result = null;
                    errorMsg = e.Message;
                }
                finally
                {
                    msIn.Close();
                    if (msOut != null)
                        msOut.Close();
                }
            }
            return result;
        }





        // Compresses the supplied memory stream, naming it as zipEntryName, into a zip,
        // which is returned as a memory stream or a byte array.
        //
        private static MemoryStream CreateToMemoryStream(MemoryStream memStreamIn, string zipEntryName, int level)
        {
            MemoryStream outputMemStream = new MemoryStream();
            ZipOutputStream zipStream = new ZipOutputStream(outputMemStream);

            zipStream.SetLevel(level); //0-9, 9 being the highest level of compression

            ZipEntry newEntry = new ZipEntry(zipEntryName);
            newEntry.DateTime = DateTime.Now;

            zipStream.PutNextEntry(newEntry);

            StreamUtils.Copy(memStreamIn, zipStream, new byte[4096]);
            zipStream.CloseEntry();

            zipStream.IsStreamOwner = false;    // False stops the Close also Closing the underlying stream.
            zipStream.Close();          // Must finish the ZipOutputStream before using outputMemStream.

            outputMemStream.Position = 0;
            return outputMemStream;
        }

        private static MemoryStream UnzipFromStream(Stream zipStream)
        {
            ZipInputStream zipInputStream = new ZipInputStream(zipStream);
            ZipEntry zipEntry = zipInputStream.GetNextEntry();
            while (zipEntry != null)
            {
                String entryFileName = zipEntry.Name;
                // to remove the folder from the entry:- entryFileName = Path.GetFileName(entryFileName);
                // Optionally match entrynames against a selection list here to skip as desired.
                // The unpacked length is available in the zipEntry.Size property.

                byte[] buffer = new byte[4096];     // 4K is optimum

                if (entryFileName.Length == 0)
                {
                    zipEntry = zipInputStream.GetNextEntry();
                    continue;
                }

                // Unzip file in buffered chunks. This is just as fast as unpacking to a buffer the full size
                // of the file, but does not waste memory.
                // The "using" will close the stream even if an exception occurs.
                //			using (FileStream streamWriter = File.Create(fullZipToPath))
                //			{
                //				StreamUtils.Copy(zipInputStream, streamWriter, buffer);
                //			}

                MemoryStream outStream = new MemoryStream();
                StreamUtils.Copy(zipInputStream, outStream, buffer);
                return outStream;

                //				zipEntry = zipInputStream.GetNextEntry();
            }
            return null;
        }





        private static MemoryStream gZip(MemoryStream memStreamIn, int level)
        {
            MemoryStream outStream = new MemoryStream();

            GZipOutputStream s = new GZipOutputStream(outStream);
            s.SetLevel(level);

            int size;
            byte[] buf = new byte[4096];
            do
            {
                size = memStreamIn.Read(buf, 0, buf.Length);
                s.Write(buf, 0, size);
            } while (size > 0);

            s.IsStreamOwner = false;
            s.Close();

            outStream.Position = 0;
            return outStream;
        }


        private static MemoryStream gunZip(MemoryStream memStreamIn)
        {
            MemoryStream outStream = new MemoryStream();

            GZipInputStream gZipStream = new GZipInputStream(memStreamIn);

            byte[] buf = new byte[4096];
            StreamUtils.Copy(gZipStream, outStream, buf);

            gZipStream.Close();

            outStream.Position = 0;
            return outStream;
        }



    }
}
