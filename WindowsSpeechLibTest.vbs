Set sapi = CreateObject("SAPI.SpVoice")
Set cat  = CreateObject("SAPI.SpObjectTokenCategory")
cat.SetID "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech_OneCore\Voices", False
For Each token In cat.EnumerateTokens
    Set oldv = sapi.Voice
    Set sapi.Voice = token
    sapi.Speak "����ɂ��́B" & token.GetAttribute("Name") & "�ł��B"
    Set sapi.Voice = oldv
Next
