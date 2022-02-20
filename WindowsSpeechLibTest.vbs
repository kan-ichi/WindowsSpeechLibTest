Set sapi = CreateObject("SAPI.SpVoice")
Set cat  = CreateObject("SAPI.SpObjectTokenCategory")
cat.SetID "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech_OneCore\Voices", False
For Each token In cat.EnumerateTokens
    Set oldv = sapi.Voice
    Set sapi.Voice = token
    sapi.Speak "Ç±ÇÒÇ…ÇøÇÕÅB" & token.GetAttribute("Name") & "Ç≈Ç∑ÅB"
    Set sapi.Voice = oldv
Next
