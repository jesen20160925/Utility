using Jesen.Utility.Encrypt;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Jesen.Utility
{
    class Program1
    {
        static void Main(string[] args)
        {
            try
            {
                #region MD5
                //1 防止看到明文 数据库密码，加盐(原密码+固定字符串，然后再MD5/双MD5)
                //2 防篡改   
                //急速秒传(第一次上传文件，保存md5摘要，第二次上传检查md5摘要) 
                //文件下载(防篡改，官方发布的时候给一个md5摘要，安装的时候首先检查下摘要)
                //svn  TFS  git  VSS(本地保存文件的md5摘要，任何修改都会影响md5)
                //3 防止抵赖

                Console.WriteLine(MD5Encrypt.Encrypt("1"));
                Console.WriteLine(MD5Encrypt.Encrypt("1"));
                Console.WriteLine(MD5Encrypt.Encrypt("123456小李"));
                Console.WriteLine(MD5Encrypt.Encrypt("113456小李"));
                Console.WriteLine(MD5Encrypt.Encrypt("113456小李113456小李113456小李113456小李113456小李113456小李113456小李"));
                string md5Abstract1 = MD5Encrypt.AbstractFile(@"D:\ruanmou\online9\homework\1\Advanced9第一次作业优秀合集.rar");
                string md5Abstract2 = MD5Encrypt.AbstractFile(@"D:\ruanmou\online9\homework\1\Advanced9第一次作业优秀合集 - 副本.rar");

                #endregion

                #region Des
                //可逆对称加密
                string desEn = DesEncrypt.Encrypt("王殃殃");
                string desDe = DesEncrypt.Decrypt(desEn);
                string desEn1 = DesEncrypt.Encrypt("张三李四");
                string desDe1 = DesEncrypt.Decrypt(desEn1);
                #endregion

                #region Rsa
                //可逆非对称加密
                KeyValuePair<string, string> encryptDecrypt = RsaEncrypt.GetKeyPair();
                string rsaEn1 = RsaEncrypt.Encrypt("net", encryptDecrypt.Key);//key是加密的
                string rsaDe1 = RsaEncrypt.Decrypt(rsaEn1, encryptDecrypt.Value);//value 解密的   不能反过来用的
                //加密钥  解密钥  钥匙的功能划分
                //公钥    私钥      公开程度划分
                #endregion

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            Console.Read();
        }
    }
}
