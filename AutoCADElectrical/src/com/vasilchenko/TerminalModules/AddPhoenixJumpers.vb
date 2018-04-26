
Imports AutoCADElectrical.com.vasilchenko.DBAccessConnection
Imports AutoCADElectrical.com.vasilchenko.TerminalClasses
Imports AutoCADElectrical.com.vasilchenko.TerminalEnums

Namespace com.vasilchenko.TerminalModules
    Module AddPhoenixJumpers
        Public Sub FillStripByJumpers(ByRef objTerminalStripList As TerminalStripClass)
            Dim objPrevItem As JumperClass = Nothing

            objTerminalStripList.JumperList = DBDataAccessObject.FillJumperData(objTerminalStripList.TerminalList.Item(0).P_TAGSTRIP)

            For Each objCurItem As JumperClass In objTerminalStripList.JumperList
                If objPrevItem Is Nothing Then
                    objCurItem.Side = SideEnum.Rigth
                ElseIf ((objPrevItem.StartTermNum + objPrevItem.TermCount - 1) = objCurItem.StartTermNum) Then
                    If objPrevItem.Side = SideEnum.Rigth Then
                        objCurItem.Side = SideEnum.Left
                    Else
                        objCurItem.Side = SideEnum.Rigth
                    End If
                Else
                        objCurItem.Side = SideEnum.Rigth
                End If
                objPrevItem = objCurItem
            Next

        End Sub

        'херим, читаем с базы данных

        '    If objTerminalStripList.TerminalList.Count = 0 Then Exit Sub
        '    Dim objCurTerminal As TerminalClass = Nothing
        '    Dim objNextTerminal As TerminalClass = Nothing
        '    Dim objPinList As New List(Of TempJumpers)

        '    If objTerminalStripList.TerminalList.Count < 2 Then Exit Sub

        '    For index As Integer = 0 To objTerminalStripList.TerminalList.Count - 2

        '        objCurTerminal = objTerminalStripList.TerminalList.Item(index)
        '        objNextTerminal = objTerminalStripList.TerminalList.Item(index + 1)

        '        ' Тут новая херовая херня :(
        '        If objCurTerminal.WiresLeftList.Count <> 0 Then
        '            For intI As Integer = 0 To objCurTerminal.WiresLeftList.Count - 1
        '                If objCurTerminal.WiresLeftList.Count <> 0 AndAlso objNextTerminal.WiresLeftList.Count <> 0 Then
        '                    AddPins(objCurTerminal, objNextTerminal, objPinList, objCurTerminal.WiresLeftList, objNextTerminal.WiresLeftList, intI)
        '                ElseIf objCurTerminal.WiresLeftList.Count <> 0 AndAlso objNextTerminal.WiresRigthList.Count <> 0 Then
        '                    AddPins(objCurTerminal, objNextTerminal, objPinList, objCurTerminal.WiresLeftList, objNextTerminal.WiresRigthList, intI)
        '                End If
        '            Next
        '        ElseIf objCurTerminal.WiresRigthList.Count <> 0 Then
        '            For intI As Integer = 0 To objCurTerminal.WiresRigthList.Count - 1
        '                If objCurTerminal.WiresRigthList.Count <> 0 AndAlso objNextTerminal.WiresLeftList.Count <> 0 Then
        '                    AddPins(objCurTerminal, objNextTerminal, objPinList, objCurTerminal.WiresRigthList, objNextTerminal.WiresLeftList, intI)
        '                ElseIf objCurTerminal.WiresRigthList.Count <> 0 AndAlso objNextTerminal.WiresRigthList.Count <> 0 Then
        '                    AddPins(objCurTerminal, objNextTerminal, objPinList, objCurTerminal.WiresRigthList, objNextTerminal.WiresRigthList, intI)
        '                End If
        '            Next
        '        End If
        '    Next

        '    Dim objConstJumperList As New List(Of KeyValuePair(Of Integer, AccessoriesPhoenixContactEnum))({
        '        New KeyValuePair(Of Integer, AccessoriesPhoenixContactEnum)(10, AccessoriesPhoenixContactEnum.Jumper10),
        '        New KeyValuePair(Of Integer, AccessoriesPhoenixContactEnum)(4, AccessoriesPhoenixContactEnum.Jumper4),
        '        New KeyValuePair(Of Integer, AccessoriesPhoenixContactEnum)(3, AccessoriesPhoenixContactEnum.Jumper3),
        '        New KeyValuePair(Of Integer, AccessoriesPhoenixContactEnum)(2, AccessoriesPhoenixContactEnum.Jumper2)
        '                                                                                              })
        '    For Each objCurItem As TempJumpers In objPinList
        '        Do While objCurItem.List.Count <> 1
        '            FillStrip(objTerminalStripList, objCurItem, objConstJumperList)
        '        Loop
        '    Next

        'End Sub

        'Private Sub FillStrip(objTerminalStripList As TerminalStripClass, objInputJumperList As TempJumpers,
        '                      objConstJumperList As List(Of KeyValuePair(Of Integer, AccessoriesPhoenixContactEnum)))

        '    Dim objJumper As New JumperClass
        '    Dim blnCheck As Boolean = False

        '    For intI = 0 To objConstJumperList.Count - 1
        '        If objInputJumperList.List.Count - objConstJumperList.Item(intI).Key > -1 Then
        '            objJumper.PinCount = objConstJumperList.Item(intI).Key
        '            objJumper.StartTermNum = objInputJumperList.List.Item(0)
        '            objJumper.Jumper = EnumFunctions.Convert(objConstJumperList.Item(intI).Value, objTerminalStripList.TerminalList.Item(0))

        '            If objTerminalStripList.JumperList Is Nothing Then
        '                objTerminalStripList.JumperList = New List(Of JumperClass)
        '                objJumper.Side = SideEnum.Rigth
        '                blnCheck = True
        '            Else
        '                For intA As Integer = 0 To objTerminalStripList.JumperList.Count - 1
        '                    Dim objTempJump As JumperClass = objTerminalStripList.JumperList.Item(intA)
        '                    If objTempJump.StartTermNum + objTempJump.PinCount - 1 = objInputJumperList.List(0) Then
        '                        If objTempJump.Side = SideEnum.Left Then
        '                            objJumper.Side = SideEnum.Rigth
        '                        Else
        '                            objJumper.Side = SideEnum.Left
        '                        End If
        '                        blnCheck = True
        '                    End If
        '                Next
        '            End If

        '            If Not blnCheck Then objJumper.Side = SideEnum.Rigth

        '            objTerminalStripList.JumperList.Add(objJumper)

        '            objInputJumperList.List.RemoveRange(0, objJumper.PinCount - 1)
        '            Exit Sub
        '        End If
        '    Next
        'End Sub

        'Private Sub AddPins(objCurTerminal As TerminalClass, objNextTerminal As TerminalClass, ByRef objPinList As List(Of TempJumpers),
        '                    objCurWiresArray As ArrayList, objNextWiresArray As ArrayList, intI As Integer)

        '    If intI > objCurWiresArray.Count - 1 Then Exit Sub

        '    For intA As Integer = 0 To objNextWiresArray.Count - 1
        '        If objCurWiresArray.Item(intI).Wireno = objNextWiresArray.Item(intA).Wireno AndAlso
        '                 objCurWiresArray.Item(intI).Pin = objNextWiresArray.Item(intA).Pin - 1 Then
        '            Dim objTempJump = objPinList.Find(Function(x) x.Wireno = objCurWiresArray.Item(intI).Wireno AndAlso
        '                                                  x.List.Item(x.List.Count - 1) = objCurWiresArray.Item(intI).PIN + 1)
        '            If objTempJump Is Nothing Then
        '                objPinList.Add(New TempJumpers(objCurWiresArray.Item(intI).Wireno,
        '                                               New List(Of Integer)({objCurTerminal.TERM, objNextTerminal.TERM})))
        '            Else objTempJump.List.Add(objNextTerminal.TERM)
        '                'objPinList.Find(Function(x)
        '                '                    If x.Wireno = objCurWiresArray.Item(intI).Wireno Then
        '                '                        x.List.Add(objNextTerminal.TERM)
        '                '                    End If
        '                '                End Function)
        '            End If
        '            'Удалить ссылки на друг-друга
        '            objCurWiresArray.RemoveAt(intI)
        '            objNextWiresArray.RemoveAt(intA)
        '            Exit Sub
        '        End If
        '    Next
        'End Sub

        'Private Class TempJumpers
        '    Dim _wireno As String
        '    Dim _list As List(Of Integer)

        '    Public Sub New()

        '    End Sub
        '    Public Sub New(wireno As String, list As List(Of Integer))
        '        _wireno = wireno
        '        _list = list
        '    End Sub

        '    Public Property Wireno As String
        '        Get
        '            Return _wireno
        '        End Get
        '        Set(value As String)
        '            Me._wireno = value
        '        End Set
        '    End Property

        '    Public Property List As List(Of Integer)
        '        Get
        '            Return _list
        '        End Get
        '        Set(value As List(Of Integer))
        '            Me._list = value
        '        End Set
        '    End Property
        'End Class

    End Module
End Namespace