
Imports SHDocVw

Public Class browserinterface
    Private m_IE As InternetExplorer
    Private m_defaultietimeout As Double

    Public Overridable Sub Launch()
        m_IE = New InternetExplorer()
        'Hide Internet Explorer Dialog Boxes.
        m_IE.Silent = True
        m_IE.Visible = True
    End Sub

    Public Overridable Sub IE_Kill()
        Try
            If m_IE IsNot Nothing Then
                m_IE.Quit()
                m_IE = Nothing
            End If
        Catch ex As Exception
        End Try
    End Sub

    Public Overridable Sub IE_GoBack()
        Try
            If m_IE IsNot Nothing Then
                m_IE.GoBack()
            End If
        Catch ex As Exception
        End Try
    End Sub

    Public Overridable Sub IE_OpenURL(ByVal v_url As String)
        m_IE.Navigate(v_url)
        IE_WaitForLoad()
        If Len(HTMLRoot.innerText) > 0 Then
            If HTMLRoot.innerText.Contains("server not responding") _
                     OrElse HTMLRoot.innerText.Contains("The page you are looking for is currently unavailable") _
                     OrElse HTMLRoot.innerText.Contains("The page might be temporarily unavailable") _
                     OrElse HTMLRoot.innerText.Contains("HTTP 404 Not Found") _
                     OrElse HTMLRoot.innerText.Contains("Internet Explorer was unable to link to the Web page") Then
                IE_Kill()
                Wait(10)
                IE_Kill()
                'exception handling
            End If
        End If
    End Sub

    Public Overridable Sub IE_WaitForLoad(Optional ByVal v_timeoutsec As Double = -1)
        Try
            Dim waitcount As Double
            Dim tStr As String = ""

            waitcount = 0
            If v_timeoutsec = -1 Then v_timeoutsec = m_defaultietimeout
            v_timeoutsec = v_timeoutsec * 10 'adjust because Wait() is called 10 times a second
            Do Until m_IE IsNot Nothing And m_IE.ReadyState = tagREADYSTATE.READYSTATE_COMPLETE
                Wait(0.1)
                waitcount = waitcount + 1
                If waitcount >= v_timeoutsec Then
                    'exception handling
                    Exit Do
                End If
                'iScraper.MessageStatus(Now.TimeOfDay.ToString + "=" + m_IE.Document.ReadyState.ToString)
            Loop
            '**********************
            Do Until m_IE IsNot Nothing And m_IE.Document.ReadyState.ToString = "complete"
                Wait(0.1)
                waitcount = waitcount + 1
                If waitcount >= v_timeoutsec Then
                    'exception handling
                    Exit Do
                End If
                'iScraper.MessageStatus(Now.TimeOfDay.ToString + "_=" + m_IE.Document.ReadyState.ToString + "]" + Asc(m_IE.Document.ReadyState.ToString) + "]")
            Loop
            'If Len(HTMLRoot.innerText) > 0 Then
            '    If HTMLRoot.innerText.Contains("server not responding") _
            '            OrElse HTMLRoot.innerText.Contains("The page you are looking for is currently unavailable") _
            '            OrElse HTMLRoot.innerText.Contains("The page might be temporarily unavailable") _
            '            OrElse HTMLRoot.innerText.Contains("HTTP 404 Not Found") _
            '            OrElse HTMLRoot.innerText.Contains("Internet Explorer was unable to link to the Web page") Then
            '        IE_Kill()
            '        Wait(10)
            '        IE_Kill()
            '        'Exception(handling)
            '    End If
            'End If
        Catch ex As Exception
            'iScraper.MessageStatus(Now.TimeOfDay.ToString + "_=" + m_IE.Document.ReadyState.ToString + "]" + Asc(m_IE.Document.ReadyState.ToString) + "]")
            'MsgBox(ex.ToString)
        End Try
    End Sub

    Public Overridable Function IE_WaitForText(ByVal tText As String, Optional ByVal v_timeoutsec As Integer = 1) As Boolean
        '***********************************************************************************************************
        '* 2012-04-23 RFK: wait for the login page to be at the correct location (their login returns with an error)
        Dim iSeconds As Integer = v_timeoutsec
        Do While iSeconds > 0
            iScraper.MessageStatus(iSeconds.ToString)
            If (HTMLRoot IsNot Nothing) Then
                If (HTMLRoot.innerText.ToString.Contains(tText)) Then Return True
            Else
                iScraper.MessageStatus("IE_WaitForText Failed")
            End If
            Wait(1)
            iSeconds = iSeconds - 1
        Loop
        iScraper.MessageStatus("Loop:" + iSeconds.ToString)
        If HTMLRoot IsNot Nothing And HTMLRoot.innerText.ToString.Contains(tText) Then Return True
        Return False
    End Function

    Public Function FindElementCustom(ByVal v_searchstring As String, ByVal v_searchtype As String, ByVal v_attributename As String, ByVal v_attributevalue As String) As mshtml.IHTMLDOMNode
        Dim elements As mshtml.IHTMLElementCollection
        Dim element As mshtml.IHTMLElement
        Dim i As Integer

        elements = FindElements(v_searchstring, v_searchtype)
        element = Nothing

        If elements.length > 1 Then
            For i = 0 To elements.length - 1
                If elements.item(i).attributes(v_attributename) IsNot Nothing AndAlso elements.item(i).attributes(v_attributename).value.ToString = v_attributevalue Then
                    element = elements.item(i)
                    Exit For
                End If
            Next
        Else
            element = elements.item(0)
        End If

        Return CType(element, mshtml.IHTMLDOMNode)

    End Function

    Public Function FindElementByID(ByVal v_searchstring As String, ByVal v_searchtype As String, ByVal v_id As String) As mshtml.IHTMLDOMNode
        Dim elements As mshtml.IHTMLElementCollection
        Dim element As mshtml.IHTMLElement
        Dim i As Integer

        elements = FindElements(v_searchstring, v_searchtype)
        element = Nothing

        If elements Is Nothing Then Return Nothing

        For i = 0 To elements.length - 1
            If elements.item(i).id = v_id Then
                element = elements.item(index:=i)
                Exit For
            End If
        Next

        Return CType(element, mshtml.IHTMLDOMNode)

    End Function

    Public Function FindElementCollection(ByVal v_searchstring As String, ByVal v_searchtype As String) As Collection
        Dim dom As mshtml.HTMLDocument
        Dim elements As Collection

        dom = m_IE.Document

        Select Case v_searchtype
            Case "CLASS"
                elements = GetElementsByClassName(dom, v_searchstring)
            Case "TEXT"
                elements = GetElementsByInnerText(dom, v_searchstring)
            Case Else
                elements = Nothing
        End Select

        Return elements
    End Function

    Public Function FindNearestTable(ByVal v_searchstring As String) As mshtml.IHTMLTable
        Dim dom As mshtml.HTMLDocument
        Dim elements As Collection
        Dim element As mshtml.IHTMLElement
        Dim table As mshtml.IHTMLTable
        dom = m_IE.Document
        elements = FindElementCollection(v_searchstring, "TEXT")

        If elements.Count > 0 AndAlso elements.Item(1) IsNot Nothing Then
            element = elements.Item(1)
            While element IsNot Nothing AndAlso element.tagName <> "TABLE"
                element = element.parentElement
            End While
            table = CType(element, mshtml.IHTMLTable)
            Return table
        Else
            Return Nothing
        End If
    End Function

    Public Sub ClickLink(ByVal v_link As mshtml.HTMLAnchorElement)
        ClickObject(v_link)
    End Sub

    Public Sub ClickImage(ByVal v_imagesrc As String)
        Dim elements As mshtml.IHTMLElementCollection
        Dim element

        elements = FindElements("IMAGE", "TAG")

        For Each element In elements
            If element.src = v_imagesrc Then
                ClickObject(element)
                Exit For
            End If
        Next
    End Sub

    Public Sub ClickObject(ByVal v_object As Object)
        v_object.click()
        'IE_WaitForLoad()
    End Sub

    Public Function ClickButton(ByVal v_btnname As String, ByVal v_btnvalue As String, Optional ByVal v_formname As String = "") As Boolean
        Dim elements As mshtml.IHTMLElementCollection
        Dim element

        elements = FindElements("INPUT", "TAG")

        For Each element In elements
            If element.name = v_btnname And element.value = v_btnvalue Then
                ClickObject(element)
                Exit For
            End If
        Next
        Return True
    End Function

    Public Sub PopulateInput(ByVal v_fieldname As String, ByVal v_fieldvalue As String, Optional ByVal v_formname As String = "")
        Dim fields As mshtml.IHTMLElementCollection
        Dim field

        Try
            fields = FindElements("INPUT", "TAG")

            If fields.item(name:=v_fieldname) IsNot Nothing Then
                field = fields.item(name:=v_fieldname)
                field.value = ""
                field.value = v_fieldvalue
            End If
        Catch ex As System.Exception

        End Try
    End Sub

    Public Sub PopulateInputByIndex(ByVal v_fieldname As String, ByVal v_fieldvalue As String, Optional ByVal v_formname As String = "", Optional ByVal v_index As Integer = -1)
        Dim fields As mshtml.IHTMLElementCollection
        Dim field As mshtml.HTMLInputElement
        Dim name As String

        Try
            fields = FindElements("INPUT", "TAG")

            If v_index <> -1 Then
                field = fields.item(index:=v_index)
                field.value = v_fieldvalue
            Else
                For Each field In fields
                    name = field.name.ToString
                    If name = v_fieldname Then
                        field.value = v_fieldvalue
                        Exit Sub
                    End If
                Next
            End If
        Catch ex As System.Exception
            Throw ex
        End Try
    End Sub

    Public Function GetInputValue(ByVal v_fieldname As String, Optional ByVal v_formname As String = "") As String
        Dim fields As mshtml.IHTMLElementCollection
        Dim field
        Dim value As String = ""

        Try
            fields = FindElements("INPUT", "TAG")

            If fields.item(name:=v_fieldname) IsNot Nothing Then
                field = fields.item(name:=v_fieldname)
                value = field.value
            End If
            Return value
        Catch ex As System.Exception
            Return value
        End Try
    End Function

    Public Function FocusOrBlur(ByVal v_fieldname As String, ByVal v_type As String, Optional ByVal v_formname As String = "") As Boolean
        Dim fields As mshtml.IHTMLElementCollection
        Dim field

        Try

            fields = FindElements("INPUT", "TAG")

            If fields.item(name:=v_fieldname) IsNot Nothing Then
                field = fields.item(name:=v_fieldname)

                If v_type = "blur" Then
                    field.blur()
                Else
                    field.focus()
                End If

            End If
        Catch ex As System.Exception

        End Try
    End Function

    Private Function GetElementsByInnerText(ByVal v_containernode, ByVal v_innertext) As Collection
        Dim elements As Collection
        Dim alltags As mshtml.IHTMLElementCollection
        Dim htmlelement

        elements = New Collection
        alltags = v_containernode.getElementsByTagName("*")

        For Each htmlelement In alltags
            If htmlelement.innerText IsNot Nothing AndAlso htmlelement.innerText.ToString.Trim = v_innertext Then
                elements.Add(htmlelement)
                Exit For
            End If
        Next
        Return elements
    End Function

    Public Sub IE_SetSelectBox(ByVal v_fieldname As String, ByVal v_fieldvalue As String)
        Dim selectboxes As mshtml.IHTMLElementCollection
        Dim selectbox As mshtml.IHTMLSelectElement
        Dim optionelement As mshtml.IHTMLOptionElement

        selectboxes = FindElements(v_fieldname, "NAME")
        For Each selectbox In selectboxes
            For Each optionelement In selectbox.options
                If optionelement.text = v_fieldvalue Then
                    optionelement.selected = True
                    Exit Sub
                End If
            Next
        Next
    End Sub

    Public Sub IE_FireEvent(ByVal fElement As String, ByVal fEvent As String)
        Dim selectbox As mshtml.HTMLSelectElement
        selectbox = Dom.getElementById(fElement)
        selectbox.FireEvent(fEvent)
    End Sub

    Public Sub IE_FireNameEvent(ByVal fElement As String, ByVal fEvent As String)
        Dim mElement As mshtml.HTMLInputElement

        mElement = Dom.getElementsByid(fElement)
        MsgBox(mElement.value.ToString)
        'mElement.FireEvent(fEvent)
    End Sub

    Public Function IE_ClickSelectBoxReturnValue(ByVal v_fieldname As String, ByVal v_fieldvalue As String) As String
        Dim selectboxes As mshtml.IHTMLElementCollection
        Dim selectbox As mshtml.IHTMLSelectElement
        Dim optionelement As mshtml.IHTMLOptionElement

        selectboxes = FindElements(v_fieldname, "NAME")
        For Each selectbox In selectboxes
            For Each optionelement In selectbox.options
                If optionelement.text = v_fieldvalue Then
                    optionelement.selected = True
                    optionelement.click()
                    IE_ClickSelectBoxReturnValue = optionelement.value
                    Exit Function
                End If
            Next
        Next
        IE_ClickSelectBoxReturnValue = Nothing
    End Function

    Private Function AscendDOM(ByVal v_element, ByVal v_target)

        While (LCase(v_element.nodeName) <> LCase(v_target) And LCase(v_element.nodeName) <> "html")
            v_element = v_element.ParentNode
        End While
        If LCase(v_element.nodeName) = "html" Then
            Return Nothing
        Else
            Return v_element
        End If
    End Function

    Public Sub FindAndClick(ByVal v_linktext As String)
        Dim links As mshtml.IHTMLElementCollection
        Dim link As mshtml.HTMLAnchorElement
        Dim j As Integer

        links = FindElements("A", "TAG")
        For j = links.length - 1 To 0 Step -1
            link = links.item(j)
            If link.href.Contains(v_linktext) Then
                ClickLink(link)
                Exit For
            End If
        Next
    End Sub

    Public ReadOnly Property Dom() As mshtml.HTMLDocument
        Get
            Dom = m_IE.Document
        End Get
    End Property

    Public Property IE() As InternetExplorer
        Get
            IE = m_IE
        End Get
        Set(ByVal v_IE As InternetExplorer)
            m_IE = Nothing
            m_IE = v_IE
        End Set
    End Property

    Public Sub New()
        m_defaultietimeout = 30 'in seconds
    End Sub

    Public Sub IE_ClickNAME(ByVal v_IDname As String)
        Dim anchors As mshtml.IHTMLElementCollection
        Dim anchor

        anchors = FindElements(v_IDname, "NAME")
        For Each anchor In anchors
            ClickObject(anchor)
            Exit For
        Next
    End Sub

    Private Function GetElementsByClassName(ByVal v_domnode As mshtml.HTMLDocument, ByVal v_classname As String) As Collection
        Dim elements As Collection
        Dim alltags As mshtml.IHTMLElementCollection
        Dim htmlelement

        alltags = v_domnode.getElementsByTagName("*")
        elements = New Collection
        For Each htmlelement In alltags
            If htmlelement.className = v_classname Then
                elements.Add(htmlelement)
                Exit For
            End If
        Next
        Return elements
    End Function

    Public Sub IE_findCLASSclick(ByVal v_searchstring As String)
        Dim dom As mshtml.HTMLDocument
        Dim elements As Collection
        Dim element As mshtml.IHTMLElement

        dom = m_IE.Document
        elements = GetElementsByClassName(dom, v_searchstring)
        For Each element In elements
            element.click()
        Next
    End Sub

    Public ReadOnly Property HTMLRoot() As mshtml.HTMLHtmlElement
        Get
            Dim htmlroots As mshtml.IHTMLElementCollection
            htmlroots = FindElements("HTML", "TAG")
            If htmlroots IsNot Nothing Then
                HTMLRoot = htmlroots.item(0)
            Else
                HTMLRoot = Nothing
            End If

        End Get
    End Property

    Public Function IE_ClickLink(ByVal v_linktext As String) As Boolean
        Dim anchors As mshtml.IHTMLElementCollection
        Dim anchor

        anchors = FindElements("A", "TAG")
        For Each anchor In anchors
            If anchor.innerHtml.ToString.Trim = v_linktext Then
                ClickObject(anchor)
                Exit For
            End If
        Next
        Return False
    End Function

    Public Function FindElements(ByVal v_searchstring As String, ByVal v_searchtype As String) As mshtml.IHTMLElementCollection
        Dim dom As mshtml.HTMLDocument
        Dim elements As mshtml.IHTMLElementCollection
        Try

            If m_IE.Document IsNot Nothing Then
                dom = m_IE.Document
                Select Case v_searchtype
                    Case "TAG"
                        'returns an mshtml.IHTMLElementCollection
                        elements = dom.getElementsByTagName(v_searchstring)
                    Case "CLASS"
                        elements = GetElementsByClassName(dom, v_searchstring)
                    Case "NAME"
                        elements = dom.getElementsByName(v_searchstring)
                    Case "ID"
                        elements = dom.getElementById(v_searchstring)
                    Case "TEXT"
                        'elements = GetElementsByInnerText(dom, v_searchstring)
                        elements = GetElementsByInnerText(dom, v_searchstring)
                    Case Else
                        elements = Nothing
                End Select
            End If
            Return elements
        Catch ex As Exception
            Return elements
        End Try
    End Function

    Public Function FindElement(ByVal v_searchstring As String, ByVal v_searchtype As String, Optional ByVal v_name As String = "") As mshtml.IHTMLDOMNode
        Dim elements As mshtml.IHTMLElementCollection
        Dim element As mshtml.IHTMLElement
        Dim i As Integer

        elements = FindElements(v_searchstring, v_searchtype)
        element = Nothing

        If elements.length > 1 And v_name <> "" Then
            For i = 0 To elements.length - 1
                If elements.item(i).name = v_name Then
                    element = elements.item(index:=i)
                    Exit For
                End If
            Next
        Else
            element = elements.item(0)
        End If

        Return CType(element, mshtml.IHTMLDOMNode)

    End Function

    Public Function IE_ClickCheckBox(ByVal v_linktext As String) As Boolean
        Dim anchors As mshtml.IHTMLElementCollection
        Dim anchor

        anchors = FindElements(v_linktext, "ID")
        For Each anchor In anchors
            'MsgBox(anchor.innerHtml.ToString.Trim)
            If anchor.innerHtml.ToString.Trim = v_linktext Then
                ClickObject(anchor)
                Exit For
            End If
        Next
        MsgBox("NO")
        Return False
    End Function

    Public Function IE_SetRadio(ByVal v_name As String, ByVal v_text As String) As Boolean
        Dim anchors As mshtml.IHTMLElementCollection
        Dim anchor As mshtml.IHTMLInputElement

        anchors = FindElements(v_name, "NAME")
        For Each anchor In anchors
            If anchor.value.ToString = v_text Then
                ClickObject(anchor)
                Return True
            End If
        Next
        Return False
    End Function

    Public Function IE_FindTextNumerousClick(ByVal v_text As String, ByVal iWhichOneToClick As Integer) As Boolean
        Dim elements As Collection
        Dim alltags As mshtml.IHTMLElementCollection
        Dim htmlelement
        Dim iFound

        iFound = 0

        elements = New Collection
        alltags = Dom.getElementsByTagName("*")
        For Each htmlelement In alltags
            If htmlelement.innerText IsNot Nothing AndAlso htmlelement.innerText.ToString.Trim = v_text Then
                iFound = iFound + 1
                If iFound = iWhichOneToClick Then
                    htmlelement.click()
                    Return True
                Else
                    'MsgBox(htmlelement.innerText.ToString)
                End If
            End If
        Next
        Return False
    End Function

End Class
