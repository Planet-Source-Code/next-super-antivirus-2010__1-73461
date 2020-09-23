VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHuffman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Huffman Compression Algorithm
' David Midkiff (mdj2023@hotmail.com>

Private Const PROGRESS_CALCFREQUENCY = 7
Private Const PROGRESS_CALCCRC = 5
Private Const PROGRESS_ENCODING = 88
Private Const PROGRESS_DECODING = 89
Private Const PROGRESS_CHECKCRC = 11

Event Progress(Procent As Integer)

Private Type HUFFMANTREE
  ParentNode As Integer
  RightNode As Integer
  LeftNode As Integer
  value As Integer
  Weight As Long
End Type

Private Type ByteArray
  count As Byte
  data() As Byte
End Type

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Function EncodeFile(SourceFile As String, DestFile As String) As Boolean
    Dim ByteArray() As Byte, Filenr As Integer
    
    If (Not FileExist(SourceFile)) Or (FileLen(SourceFile) = "0") Then
        EncodeFile = False
        Exit Function
    End If
    
    Filenr = FreeFile
    Open SourceFile For Binary As #Filenr
        ReDim ByteArray(0 To LOF(Filenr) - 1)
        Get #Filenr, , ByteArray()
    Close #Filenr
    
    Call EncodeByte(ByteArray(), UBound(ByteArray) + 1)
    
    If (FileExist(DestFile)) Then Kill DestFile
    
    Open DestFile For Binary As #Filenr
        Put #Filenr, , ByteArray()
    Close #Filenr
    EncodeFile = True
End Function
Public Function DecodeFile(SourceFile As String, DestFile As String) As Boolean
    Dim ByteArray() As Byte, Filenr As Integer
    
    If (Not FileExist(SourceFile)) Or (FileLen(SourceFile) = "0") Then
        DecodeFile = False
        Exit Function
    End If
    
    Filenr = FreeFile
    Open SourceFile For Binary As #Filenr
        ReDim ByteArray(0 To LOF(Filenr) - 1)
        Get #Filenr, , ByteArray()
    Close #Filenr
    
    Call DecodeByte(ByteArray(), UBound(ByteArray) + 1)
    
    If (FileExist(DestFile)) Then Kill DestFile
    
    Open DestFile For Binary As #Filenr
        Put #Filenr, , ByteArray()
    Close #Filenr
    DecodeFile = True
End Function
Private Sub CreateTree(Nodes() As HUFFMANTREE, NodesCount As Long, Char As Long, Bytes As ByteArray)
    Dim A As Integer, NodeIndex As Long
    
    NodeIndex = 0
    For A = 0 To (Bytes.count - 1)
        If (Bytes.data(A) = 0) Then
            If (Nodes(NodeIndex).LeftNode = -1) Then
                Nodes(NodeIndex).LeftNode = NodesCount
                Nodes(NodesCount).ParentNode = NodeIndex
                Nodes(NodesCount).LeftNode = -1
                Nodes(NodesCount).RightNode = -1
                Nodes(NodesCount).value = -1
                NodesCount = NodesCount + 1
            End If
            NodeIndex = Nodes(NodeIndex).LeftNode
        ElseIf (Bytes.data(A) = 1) Then
            If (Nodes(NodeIndex).RightNode = -1) Then
                Nodes(NodeIndex).RightNode = NodesCount
                Nodes(NodesCount).ParentNode = NodeIndex
                Nodes(NodesCount).LeftNode = -1
                Nodes(NodesCount).RightNode = -1
                Nodes(NodesCount).value = -1
                NodesCount = NodesCount + 1
            End If
            NodeIndex = Nodes(NodeIndex).RightNode
        Else
            Stop
        End If
    Next
    Nodes(NodeIndex).value = Char
End Sub
Public Sub EncodeByte(ByteArray() As Byte, ByteLen As Long)
    Dim I As Long, j As Long, Char As Byte, BitPos As Byte, lNode1 As Long
    Dim lNode2 As Long, lNodes As Long, lLength As Long, count As Integer
    Dim lWeight1 As Long, lWeight2 As Long, Result() As Byte, ByteValue As Byte
    Dim ResultLen As Long, Bytes As ByteArray, NodesCount As Integer, NewProgress As Integer
    Dim CurrProgress As Integer, BitValue(0 To 7) As Byte, CharCount(0 To 255) As Long
    Dim Nodes(0 To 511) As HUFFMANTREE, CharValue(0 To 255) As ByteArray
  
    If (ByteLen = 0) Then
        ReDim Preserve ByteArray(0 To ByteLen + 3)
        If (ByteLen > 0) Then Call CopyMem(ByteArray(4), ByteArray(0), ByteLen)
        ByteArray(0) = 72
        ByteArray(1) = 69
        ByteArray(2) = 48
        ByteArray(3) = 13
        Exit Sub
    End If
  
    ReDim Result(0 To 522)
    Result(0) = 72
    Result(1) = 69
    Result(2) = 51
    Result(3) = 13
    ResultLen = 4
  
    For I = 0 To (ByteLen - 1)
        CharCount(ByteArray(I)) = CharCount(ByteArray(I)) + 1
        If (I Mod 1000 = 0) Then
            NewProgress = I / ByteLen * PROGRESS_CALCFREQUENCY
            If (NewProgress <> CurrProgress) Then
                CurrProgress = NewProgress
                RaiseEvent Progress(CurrProgress)
            End If
        End If
    Next
    For I = 0 To 255
        If (CharCount(I) > 0) Then
            With Nodes(NodesCount)
                .Weight = CharCount(I)
                .value = I
                .LeftNode = -1
                .RightNode = -1
                .ParentNode = -1
            End With
            NodesCount = NodesCount + 1
        End If
    Next
  
    For lNodes = NodesCount To 2 Step -1
        lNode1 = -1: lNode2 = -1
        For I = 0 To (NodesCount - 1)
            If (Nodes(I).ParentNode = -1) Then
                If (lNode1 = -1) Then
                    lWeight1 = Nodes(I).Weight
                    lNode1 = I
                ElseIf (lNode2 = -1) Then
                    lWeight2 = Nodes(I).Weight
                    lNode2 = I
                ElseIf (Nodes(I).Weight < lWeight1) Then
                    If (Nodes(I).Weight < lWeight2) Then
                        If (lWeight1 < lWeight2) Then
                            lWeight2 = Nodes(I).Weight
                            lNode2 = I
                        Else
                            lWeight1 = Nodes(I).Weight
                            lNode1 = I
                        End If
                    Else
                        lWeight1 = Nodes(I).Weight
                        lNode1 = I
                    End If
                ElseIf (Nodes(I).Weight < lWeight2) Then
                    lWeight2 = Nodes(I).Weight
                    lNode2 = I
                End If
            End If
        Next
    
        With Nodes(NodesCount)
            .Weight = lWeight1 + lWeight2
            .LeftNode = lNode1
            .RightNode = lNode2
            .ParentNode = -1
            .value = -1
        End With
    
        Nodes(lNode1).ParentNode = NodesCount
        Nodes(lNode2).ParentNode = NodesCount
        NodesCount = NodesCount + 1
    Next

    ReDim Bytes.data(0 To 255)
    Call CreateBitSequences(Nodes(), NodesCount - 1, Bytes, CharValue)
  
    For I = 0 To 255
        If (CharCount(I) > 0) Then lLength = lLength + CharValue(I).count * CharCount(I)
    Next
    lLength = IIf(lLength Mod 8 = 0, lLength \ 8, lLength \ 8 + 1)
  
    If ((lLength = 0) Or (lLength > ByteLen)) Then
        ReDim Preserve ByteArray(0 To ByteLen + 3)
        Call CopyMem(ByteArray(4), ByteArray(0), ByteLen)
        ByteArray(0) = 72
        ByteArray(1) = 69
        ByteArray(2) = 48
        ByteArray(3) = 13
        Exit Sub
    End If
  
    Char = 0
    For I = 0 To (ByteLen - 1)
        Char = Char Xor ByteArray(I)
        If (I Mod 10000 = 0) Then
            NewProgress = I / ByteLen * PROGRESS_CALCCRC + PROGRESS_CALCFREQUENCY
            If (NewProgress <> CurrProgress) Then
                CurrProgress = NewProgress
                RaiseEvent Progress(CurrProgress)
            End If
        End If
    Next
    Result(ResultLen) = Char
    ResultLen = ResultLen + 1
    Call CopyMem(Result(ResultLen), ByteLen, 4)
    ResultLen = ResultLen + 4
    BitValue(0) = 2 ^ 0
    BitValue(1) = 2 ^ 1
    BitValue(2) = 2 ^ 2
    BitValue(3) = 2 ^ 3
    BitValue(4) = 2 ^ 4
    BitValue(5) = 2 ^ 5
    BitValue(6) = 2 ^ 6
    BitValue(7) = 2 ^ 7
    count = 0
    For I = 0 To 255
        If (CharValue(I).count > 0) Then count = count + 1
    Next
    Call CopyMem(Result(ResultLen), count, 2)
    ResultLen = ResultLen + 2
    count = 0
    For I = 0 To 255
        If (CharValue(I).count > 0) Then
            Result(ResultLen) = I
            ResultLen = ResultLen + 1
            Result(ResultLen) = CharValue(I).count
            ResultLen = ResultLen + 1
            count = count + 16 + CharValue(I).count
        End If
    Next
  
    ReDim Preserve Result(0 To ResultLen + count \ 8)
  
    BitPos = 0
    ByteValue = 0
    For I = 0 To 255
        With CharValue(I)
            If (.count > 0) Then
                For j = 0 To (.count - 1)
                    If (.data(j)) Then ByteValue = ByteValue + BitValue(BitPos)
                    BitPos = BitPos + 1
                    If (BitPos = 8) Then
                        Result(ResultLen) = ByteValue
                        ResultLen = ResultLen + 1
                        ByteValue = 0
                        BitPos = 0
                    End If
                Next
            End If
        End With
    Next
    If (BitPos > 0) Then
        Result(ResultLen) = ByteValue
        ResultLen = ResultLen + 1
    End If
  
    ReDim Preserve Result(0 To ResultLen - 1 + lLength)
  
    Char = 0
    BitPos = 0
    For I = 0 To (ByteLen - 1)
        With CharValue(ByteArray(I))
            For j = 0 To (.count - 1)
                If (.data(j) = 1) Then Char = Char + BitValue(BitPos)
                BitPos = BitPos + 1
                If (BitPos = 8) Then
                    Result(ResultLen) = Char
                    ResultLen = ResultLen + 1
                    BitPos = 0
                    Char = 0
                End If
            Next
        End With
        If (I Mod 10000 = 0) Then
            NewProgress = I / ByteLen * PROGRESS_ENCODING + PROGRESS_CALCCRC + PROGRESS_CALCFREQUENCY
            If (NewProgress <> CurrProgress) Then
                CurrProgress = NewProgress
                RaiseEvent Progress(CurrProgress)
            End If
        End If
    Next

    If (BitPos > 0) Then
        Result(ResultLen) = Char
        ResultLen = ResultLen + 1
    End If
    ReDim ByteArray(0 To ResultLen - 1)
    Call CopyMem(ByteArray(0), Result(0), ResultLen)
    If (CurrProgress <> 100) Then RaiseEvent Progress(100)
End Sub
Public Function DecodeString(Text As String) As String
    Dim ByteArray() As Byte
    ByteArray() = StrConv(Text, vbFromUnicode)
    Call DecodeByte(ByteArray, Len(Text))
    DecodeString = StrConv(ByteArray(), vbUnicode)
End Function
Public Function EncodeString(Text As String) As String
    Dim ByteArray() As Byte
    ByteArray() = StrConv(Text, vbFromUnicode)
    Call EncodeByte(ByteArray, Len(Text))
    EncodeString = StrConv(ByteArray(), vbUnicode)
End Function
Public Sub DecodeByte(ByteArray() As Byte, ByteLen As Long)
    Dim I As Long, j As Long, pos As Long, Char As Byte, CurrPos As Long
    Dim count As Integer, CheckSum As Byte, Result() As Byte, BitPos As Integer
    Dim NodeIndex As Long, ByteValue As Byte, ResultLen As Long, NodesCount As Long
    Dim lResultLen As Long, NewProgress As Integer, CurrProgress As Integer, BitValue(0 To 7) As Byte
    Dim Nodes(0 To 511) As HUFFMANTREE, CharValue(0 To 255) As ByteArray
      
    If (ByteArray(0) <> 72) Or (ByteArray(1) <> 69) Or (ByteArray(3) <> 13) Then
    ElseIf (ByteArray(2) = 48) Then
        Call CopyMem(ByteArray(0), ByteArray(4), ByteLen - 4)
        ReDim Preserve ByteArray(0 To ByteLen - 5)
        Exit Sub
    ElseIf (ByteArray(2) <> 51) Then
        Err.Raise vbObjectError, "HuffmanDecode()", "The data either was not compressed with HE3 or is corrupt (identification string not found)"
        Exit Sub
    End If
      
    CurrPos = 5
    CheckSum = ByteArray(CurrPos - 1)
    CurrPos = CurrPos + 1
      
    Call CopyMem(ResultLen, ByteArray(CurrPos - 1), 4)
    CurrPos = CurrPos + 4
    lResultLen = ResultLen
    If (ResultLen = 0) Then Exit Sub
    ReDim Result(0 To ResultLen - 1)
    Call CopyMem(count, ByteArray(CurrPos - 1), 2)
    CurrPos = CurrPos + 2
    
    For I = 1 To count
        With CharValue(ByteArray(CurrPos - 1))
            CurrPos = CurrPos + 1
            .count = ByteArray(CurrPos - 1)
            CurrPos = CurrPos + 1
            ReDim .data(0 To .count - 1)
        End With
    Next
    
    BitValue(0) = 2 ^ 0
    BitValue(1) = 2 ^ 1
    BitValue(2) = 2 ^ 2
    BitValue(3) = 2 ^ 3
    BitValue(4) = 2 ^ 4
    BitValue(5) = 2 ^ 5
    BitValue(6) = 2 ^ 6
    BitValue(7) = 2 ^ 7
    
    ByteValue = ByteArray(CurrPos - 1)
    CurrPos = CurrPos + 1
    BitPos = 0
    
    For I = 0 To 255
        With CharValue(I)
            If (.count > 0) Then
                For j = 0 To (.count - 1)
                    If (ByteValue And BitValue(BitPos)) Then .data(j) = 1
                    BitPos = BitPos + 1
                    If (BitPos = 8) Then
                        ByteValue = ByteArray(CurrPos - 1)
                        CurrPos = CurrPos + 1
                        BitPos = 0
                    End If
                Next
            End If
        End With
    Next
    
    If (BitPos = 0) Then CurrPos = CurrPos - 1
      
    NodesCount = 1
    Nodes(0).LeftNode = -1
    Nodes(0).RightNode = -1
    Nodes(0).ParentNode = -1
    Nodes(0).value = -1
    
    For I = 0 To 255
        Call CreateTree(Nodes(), NodesCount, I, CharValue(I))
    Next
      
    ResultLen = 0
    For CurrPos = CurrPos To ByteLen
        ByteValue = ByteArray(CurrPos - 1)
        For BitPos = 0 To 7
            If (ByteValue And BitValue(BitPos)) Then NodeIndex = Nodes(NodeIndex).RightNode Else NodeIndex = Nodes(NodeIndex).LeftNode
            If (Nodes(NodeIndex).value > -1) Then
                Result(ResultLen) = Nodes(NodeIndex).value
                ResultLen = ResultLen + 1
                If (ResultLen = lResultLen) Then GoTo DecodeFinished
                NodeIndex = 0
            End If
        Next
        If (CurrPos Mod 10000 = 0) Then
            NewProgress = CurrPos / ByteLen * PROGRESS_DECODING
            If (NewProgress <> CurrProgress) Then
                CurrProgress = NewProgress
                RaiseEvent Progress(CurrProgress)
            End If
        End If
    Next

DecodeFinished:
    Char = 0
    For I = 0 To (ResultLen - 1)
        Char = Char Xor Result(I)
        If (I Mod 10000 = 0) Then
            NewProgress = I / ResultLen * PROGRESS_CHECKCRC + PROGRESS_DECODING
            If (NewProgress <> CurrProgress) Then
                CurrProgress = NewProgress
                RaiseEvent Progress(CurrProgress)
            End If
        End If
    Next
    If (Char <> CheckSum) Then Err.Raise vbObjectError, "AL127 Worm Cleaner", "The data might be corrupted (checksum did not match expected value)"
    ReDim ByteArray(0 To ResultLen - 1)
    Call CopyMem(ByteArray(0), Result(0), ResultLen)
    If (CurrProgress <> 100) Then RaiseEvent Progress(100)
End Sub
Private Sub CreateBitSequences(Nodes() As HUFFMANTREE, ByVal NodeIndex As Integer, Bytes As ByteArray, CharValue() As ByteArray)
    Dim NewBytes As ByteArray
    If (Nodes(NodeIndex).value > -1) Then
        CharValue(Nodes(NodeIndex).value) = Bytes
        Exit Sub
    End If
    If (Nodes(NodeIndex).LeftNode > -1) Then
        NewBytes = Bytes
        NewBytes.data(NewBytes.count) = 0
        NewBytes.count = NewBytes.count + 1
        Call CreateBitSequences(Nodes(), Nodes(NodeIndex).LeftNode, NewBytes, CharValue)
    End If
    If (Nodes(NodeIndex).RightNode > -1) Then
        NewBytes = Bytes
        NewBytes.data(NewBytes.count) = 1
        NewBytes.count = NewBytes.count + 1
        Call CreateBitSequences(Nodes(), Nodes(NodeIndex).RightNode, NewBytes, CharValue)
    End If
End Sub

Private Function FileExist(FileName As String) As Boolean
    On Error GoTo FileDoesNotExist
    Call FileLen(FileName)
    FileExist = True
    Exit Function

FileDoesNotExist:
    FileExist = False
End Function



