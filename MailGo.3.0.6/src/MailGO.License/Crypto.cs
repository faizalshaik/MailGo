/* Copyright 2008 Data Design Vietnam. All rights reserved.
 * 
 * Created 2008.01.22 Tran Dinh Thoai
 * 
 */

using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Security.Cryptography;

namespace DataDesign.MailGO.License
{
    internal class Crypto
    {
        public static string Encrypt(string v_source, string v_key, string v_init_vec)
        {
            TripleDESCryptoServiceProvider t_provider = new TripleDESCryptoServiceProvider();
            MemoryStream t_output = new MemoryStream();

            t_provider.Mode = CipherMode.CBC;
            t_provider.Key = Encoding.ASCII.GetBytes(v_key);
            t_provider.IV = Encoding.Unicode.GetBytes(v_init_vec);

            CryptoStream t_crypto = new CryptoStream(t_output, t_provider.CreateEncryptor(), CryptoStreamMode.Write);
            Byte[] t_data = Encoding.Unicode.GetBytes(v_source);

            t_crypto.Write(t_data, 0, t_data.Length);
            t_crypto.FlushFinalBlock();
            t_crypto.Close();

            return Convert.ToBase64String(t_output.ToArray());
        }

        public static string Decrypt(string v_source, string v_key, string v_init_vec)
        {
            TripleDESCryptoServiceProvider t_provider = new TripleDESCryptoServiceProvider();
            MemoryStream t_output = new MemoryStream();

            t_provider.Mode = CipherMode.CBC;
            t_provider.Key = Encoding.ASCII.GetBytes(v_key);
            t_provider.IV = Encoding.Unicode.GetBytes(v_init_vec);

            CryptoStream t_crypto = new CryptoStream(t_output, t_provider.CreateDecryptor(), CryptoStreamMode.Write);
            Byte[] t_data = Convert.FromBase64String(v_source);

            t_crypto.Write(t_data, 0, t_data.Length);
            t_crypto.FlushFinalBlock();
            t_crypto.Close();

            return Encoding.Unicode.GetString(t_output.ToArray());
        }
    }
}
