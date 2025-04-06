
#Requires AutoHotkey v2

SetupTray() {
    I_Icon := StrReplace(A_ScriptDir "\icons\" . A_ScriptName, ".ahk", ".ico")
    if FileExist(I_Icon)
        TraySetIcon(I_Icon)

    Tray := A_TrayMenu

    links := { Github: "https://github.com/alexofrhodes/"
             , YouTube: "https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg"
             , Blog: "https://alexofrhodes.github.io/"
             , BuyMeACoffee: "https://www.buymeacoffee.com/AlexOfRhodes"
             , Gmail: "mailto:anastasioualex@gmail.com?subject=" A_ScriptName "&body=Hi! I would like to talk about ..."
            , }

    for name, url in links.ownprops() {
        Tray.Add(name, FollowLink.Bind(url))
        icon := A_ScriptDir "\icons\" name ".ico"
        Tray.SetIcon(name, icon)
    }
}

FollowLink(url,*) {
    Run url
}
