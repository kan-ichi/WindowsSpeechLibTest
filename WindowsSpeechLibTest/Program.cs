using System;

namespace WindowsSpeechLibTest
{
    class Program
    {
        static void Main(string[] args)
        {
            SapiTest.Speak();

            SapiTalk sapi = new SapiTalk();
            foreach (var item in sapi.Talkers())
            {
                Console.WriteLine("{0} {1}", item.Key, item.Value);

                sapi.SetTalker(item.Key);
                if (item.Value.Contains("Japanese"))
                {
                    sapi.Talk($"こんにちは。{ item.Value }です。");
                }
                else
                {
                    sapi.Talk($"Hellow. I am { item.Value }.");
                }

            }
        }
    }
}
