    Dim wshshell
    intOpcao = msgbox("Quer ver as noticias?",vbyesno,"NOTICIAS")
    if intOpcao = vbyes then
         Set WshShell = WScript.CreateObject("WScript.Shell")
 WshShell.Run("https://google.ru")
 WshShell.Run("https://g1.globo.com")
WshShell.Run("https://www.youtube.com/watch?v=Z-VfaG9ZN_U/?autoplay=1")
    end if
