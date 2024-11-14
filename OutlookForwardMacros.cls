Dim WithEvents InboxItems As Outlook.Items
 
' Инициализация обработчика событий при запуске Outlook
Private Sub Application_Startup()
    Dim olNs As Outlook.NameSpace
    Dim InboxFolder As Outlook.Folder
    Dim account As Outlook.account
 
    Set olNs = Application.GetNamespace("MAPI")
    ' Аккаунт почты отправителя
    Set account = olNs.Accounts("example@gmail.com")
    ' Папка "Входящие" указанного аккаунта
    On Error Resume Next
    Set InboxFolder = account.DeliveryStore.GetDefaultFolder(olFolderInbox)
    On Error GoTo 0
    If Not InboxFolder Is Nothing Then
        Set InboxItems = InboxFolder.Items
    Else
        MsgBox "Inbox folder not found in account " & account.DisplayName & ". Check the account settings."
    End If
End Sub
 
' Обработчик событий для новых сообщений в папке Входящие
Private Sub InboxItems_ItemAdd(ByVal Item As Object)
    If TypeOf Item Is MailItem Then
        Dim newMail As MailItem
        Dim senderAddress As String
        Dim tempFilePath As String
        Dim originalAttachment As Attachment
        ' Проверка, является ли отправитель нужным пользователем
        senderAddress = Item.SenderEmailAddress
        If senderAddress = "example@gmail.com" Then  ' Таргет адрес почты отправителя
            Set newMail = Application.CreateItem(olMailItem)
            With newMail
                .To = "example@gmail.com"  ' Аккаунт почты получателя
                .Subject = Item.Subject
                .Body = Item.Body
                .HTMLBody = Item.HTMLBody  ' Если нужно сохранить форматирование HTML
                ' Копируем все вложения из оригинального письма
                For Each originalAttachment In Item.Attachments
                    ' Сохраняем вложение во временное местоположение
                    tempFilePath = Environ("TEMP") & "\" & originalAttachment.FileName
                    originalAttachment.SaveAsFile tempFilePath
                    ' Добавляем вложение в новое письмо
                    .Attachments.Add tempFilePath
                Next originalAttachment
                .SendUsingAccount = Application.Session.Accounts("example@gmail.com")  ' Аккаунт почты отправителя
                .Send
            End With
            ' Удаляем временные файлы
            For Each originalAttachment In Item.Attachments
                tempFilePath = Environ("TEMP") & "\" & originalAttachment.FileName
                Kill tempFilePath
            Next originalAttachment
            ' Отмечаем письмо как непрочитанное после пересылки
            Item.UnRead = True
        End If
    End If
End Sub
