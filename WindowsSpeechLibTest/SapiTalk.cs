using SpeechLib; // 参考URL：https://qiita.com/Q11Q/items/9325cb55e8799f620fd6
using System;
using System.Collections.Generic;
using System.Threading;

namespace WindowsSpeechLibTest
{
    public class SapiTalk
    {
        private SpVoice sapi = null;
        private Dictionary<int, SpObjectToken> SpeakerList = new Dictionary<int, SpObjectToken>();

        private int Speed = 0;
        private int Volume = 100;
        private int AvatorIdx = 0;

        public SapiTalk()
        {
            int idx = 0;

            try
            {
                sapi = new SpVoice();
                SpObjectTokenCategory sapiCat = new SpObjectTokenCategory();
                Dictionary<string, SpObjectToken> TokerPool = new Dictionary<string, SpObjectToken>();

                // 参考URL：https://qiita.com/7shi/items/7781516d6746e29c03b4
                sapiCat.SetId(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech_OneCore\Voices", false);

                foreach (SpObjectToken token in sapiCat.EnumerateTokens())
                {
                    if (!TokerPool.ContainsKey(token.GetAttribute("name")))
                    {
                        TokerPool.Add(token.GetAttribute("name"), token);
                    }
                }

                foreach (SpObjectToken token in sapi.GetVoices("", ""))
                {
                    if (!TokerPool.ContainsKey(token.GetAttribute("name")))
                    {
                        TokerPool.Add(token.GetAttribute("name"), token);
                    }
                }


                foreach (var item in TokerPool)
                {
                    SpeakerList.Add(idx, item.Value);
                    idx++;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("{0},{1},{2}", e.Message, e.InnerException == null ? "" : e.InnerException.Message, e.StackTrace);
            }
        }

        public Dictionary<int, string> Talkers()
        {
            Dictionary<int, string> ans = new Dictionary<int, string>();

            for (int i = 0; i < SpeakerList.Count; i++)
            {
                ans.Add(i, SpeakerList[i].GetDescription());
            }

            return ans;
        }

        public void SetTalker(int _talker)
        {
            AvatorIdx = _talker;
        }

        public void SetVolume(int _volume)
        {
            Volume = _volume;
        }

        public void SetRate(int _speed)
        {
            Speed = _speed;
        }

        public void Talk(string _text, bool _asyncFlag = false)
        {
            try
            {
                SpObjectToken backupSapi = null;

                Thread t = new Thread(() =>
                {
                    backupSapi = sapi.Voice;
                    sapi.Voice = SpeakerList[AvatorIdx];
                    sapi.Rate = Speed;
                    sapi.Volume = Volume;
                    sapi.Speak(_text);
                    sapi.Voice = backupSapi;
                });
                t.SetApartmentState(ApartmentState.STA);
                t.Start();
                if (!_asyncFlag) t.Join();
            }
            catch
            {
                throw;
            }
        }

        public void Save(string _filePath, string _text)
        {
            try
            {
                SpObjectToken backupSapi = null;
                SpFileStream ss = new SpFileStream();
                ss.Open(_filePath, SpeechStreamFileMode.SSFMCreateForWrite);
                sapi.AudioOutputStream = ss;

                Thread t = new Thread(() => {
                    backupSapi = sapi.Voice;
                    sapi.Voice = SpeakerList[AvatorIdx];
                    sapi.Rate = Speed;
                    sapi.Volume = Volume;
                    sapi.Speak(_text);
                    sapi.Voice = backupSapi;
                });
                t.SetApartmentState(ApartmentState.STA);
                t.Start();
                t.Join();
                ss.Close();
            }
            catch
            {
                throw;
            }
        }

    }
}
