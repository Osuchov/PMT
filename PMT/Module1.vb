Module Module1

    Sub Main()
        Dim IE As Object

        IE = CreateObject("InternetExplorer.Application")

        IE.Visible = True
        IE.Navigate("http://www.onpallet.com/")

        ' Clean up
        IE.Quit
        IE = Nothing
    End Sub

End Module
