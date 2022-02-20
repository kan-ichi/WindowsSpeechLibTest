using SpeechLib;
using System;
using System.Runtime.InteropServices;

namespace WindowsSpeechLibTest
{
    static class SapiTest
    {
        public static void Speak()
        {
            SpVoice voice = new SpVoice();
            SpObjectTokenCategory otc = new SpObjectTokenCategory();
            // otc.SetId(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices"); // 「Microsoft Haruka Desktop」はこちらに存在する
            otc.SetId(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech_OneCore\Voices"); // 「Microsoft Sayaka」はこちらに存在する
            ISpeechObjectTokens tokenEnum = otc.EnumerateTokens();
            int nTokenCount = tokenEnum.Count;
            Console.WriteLine("Number of voices: {0}", nTokenCount);
            foreach (ISpeechObjectToken sot in tokenEnum)
            {
                Console.WriteLine("Voice : {0}", sot.GetDescription());
                {
                    voice.Voice = (SpObjectToken)sot;
                    try
                    {
                        if (sot.GetDescription().Contains("Japanese"))
                        {
                            voice.Speak($"初めまして。{ sot.GetDescription() }です。");
                        }
                        else
                        {
                            voice.Speak($"Hellow.I am { sot.GetDescription() }.");
                        }
                    }
                    catch (Exception)
                    {
                        Console.WriteLine("Cannot speak with the voice : {0}", sot.GetDescription());
                    }
                    Marshal.ReleaseComObject(sot);
                }
            }
        }
    }
}
