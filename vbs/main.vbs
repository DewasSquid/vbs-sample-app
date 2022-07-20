' before window loading
' create an header div
Dim headerDiv

Set headerDiv = document.CreateElement("div")
Call headerDiv.SetAttribute("id", "header")
document.body.appendChild(headerDiv)

' Add the logo to the header div
headerDiv.innerHTML = "<a href='index.hta'><img src='img/logo.png' alt='Main Page'/></a>"

Sub Window_OnLoad()
    ' get the current file name
    Dim currentFile
    
    currentFile = location.pathname ' Current file path
    currentFile = Mid(currentFile, InStrRev(currentFile, "\") + 1) ' Remove every strings followed by "/"
    currentFile = LCase(currentFile)
    currentFile = Replace(currentFile, ".hta", "") ' Remove the file extension

    ' Load each of the css files
    Dim cssFile, cssLink

    cssFiles = Array("css/fonts.css", "css/style.css", "css/style_" & currentFile & ".css")

    For Each cssFile In cssFiles
        Set cssLink = window.document.createElement("link")
        Call cssLink.setAttribute("rel", "stylesheet")
        Call cssLink.setAttribute("type", "text/css")
        Call cssLink.setAttribute("href", cssFile)
        document.appendChild(cssLink)
    Next

    ' Load each of the script files
    Dim vbsFile, vbsLink

    vbsFiles = Array("vbs/" & currentFile & ".vbs")

    For Each vbsFile In vbsFiles
        Set vbsLink = window.document.createElement("script")
        Call vbsLink.setAttribute("language", "vbscript")
        Call vbsLink.setAttribute("src", vbsFile)
        document.appendChild(vbsLink)
    Next
End Sub ' Window_OnLoad