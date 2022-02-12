VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CircularBufferByte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private allowOverwrite_ As Boolean
Private buffer_() As Byte
Private capacity_ As Long
Private head_ As Long
Private size_ As Long
Private tail_ As Long
Public Sub Create(ByVal newCapacity As Long, Optional newAllowOverwrite As Boolean = False)
    If Capacity < 0 Then
        Debug.Print "The buffer capacity must be greater than or equal to zero."
        Exit Sub
    End If
    size_ = 0
    Capacity = newCapacity
    allowOverwrite_ = newAllowOverwrite
End Sub
Public Property Get AllowOverwrite() As Boolean
    AllowOverwrite = allowOverwrite_
End Property
Public Property Let AllowOverwrite(newValue As Boolean)
    allowOverwrite_ = newValue
End Property
Public Property Get Capacity() As Long
    Capacity = capacity_
End Property
Public Property Let Capacity(newValue As Long)
    If newValue <> capacity_ Then
        Dim newBuffer() As Byte
        If newValue < size_ Then
            Debug.Print "The new capacity must be greater than or equal to the buffer size."
            Exit Property
        End If
        ReDim Preserve buffer_(newValue - 1)
        tail_ = size_
        head_ = 0
        capacity_ = newValue
    End If
End Property
Public Property Get Head() As Long
    Head = head_
End Property
Public Function IsEmpty() As Boolean
    If size_ = 0 Then
        IsEmpty = True
    Else
        IsEmpty = False
    End If
End Function
Public Function IsFull() As Boolean
    If allowOverwrite_ = False And size_ = capacity_ Then
        IsFull = True
    Else
        IsFull = False
    End If
End Function
Public Property Get Size() As Long
    Size = size_
End Property
Public Property Get Tail() As Long
    Tail = tail_
End Property
Public Sub Clear()
    size_ = 0
    head_ = 0
    tail_ = 0
    ReDim buffer_(capacity_ - 1)
End Sub
Public Function Contains(item As Byte) As Boolean
    Dim bufferIndex&, i&
    Dim result As Boolean
    bufferIndex = head_
    result = False
    For i = 0 To (size_ - 1)
        If bufferIndex = capacity_ Then
            bufferIndex = 0
        End If
        If buffer_(bufferIndex) = item Then
            result = True
            Exit For
        End If
        bufferIndex = bufferIndex + 1
    Next
    Contains = result
End Function

Public Sub CopyTo1(ByRef destArray() As Byte)
    CopyTo2 destArray, 0
End Sub
Public Sub CopyTo2(ByRef destArray() As Byte, ByVal arrayIndex As Long)
    Dim count&
    If size_ < (UBound(destArray) - arrayIndex + 1) Then
        count = size_
    Else
        count = UBound(destArray) - arrayIndex + 1
    End If
    CopyTo3 0, destArray, arrayIndex, count
End Sub
Public Sub CopyTo3(ByVal index As Long, ByRef destArray() As Byte, ByVal arrayIndex As Long, ByVal count As Long)
    Dim i&, bufferIndex&
    If count > size_ Then
        Debug.Print "The read count cannot be greater than the buffer size."
        Exit Sub
    End If
    bufferIndex = head_ + index
    
    For i = 0 To (count - 1)
        If bufferIndex >= capacity_ Then
            bufferIndex = bufferIndex - capacity_
        End If
        destArray(arrayIndex) = buffer_(bufferIndex)
        bufferIndex = bufferIndex + 1
        arrayIndex = arrayIndex + 1
    Next
End Sub

Public Function GetByteArray(ByRef destArray() As Byte, ByVal arrayIndex As Long, ByVal count As Long) As Long
    Dim realCount&, dstIndex&
    Dim i%
    If count < size_ Then
        realCount = count
    Else
        realCount = size_
    End If
    dstIndex = arrayIndex
    For i = 0 To (realCount - 1)
        If head_ = capacity_ Then
            head_ = 0
        End If
        destArray(dstIndex) = buffer_(head_)
        head_ = head_ + 1
        dstIndex = dstIndex + 1
    Next
    If head_ = capacity_ Then
        head_ = 0
    End If
    size_ = size_ - realCount
    GetByteArray = realCount
End Function
    
Public Function GetByte() As Byte
    Dim item As Byte
    If IsEmpty = True Then
        Debug.Print "The buffer is empty."
        Exit Function
    End If
    item = buffer_(head_)
    If (head_ + 1) = capacity_ Then
        head_ = 0
    End If
    size_ = size_ - 1
    GetByte = item
End Function

Public Function GetLastByte() As Byte
    Dim item As Byte
    Dim index&
    If IsEmpty = True Then
        Debug.Print "The buffer is empty."
        Exit Function
    End If
    index = GetTailIndex(0)
    item = buffer_(index)
    If (tail_ - 1) < 0 Then
        tail_ = 0
    End If
    size_ = size_ - 1
    GetLastByte = item
End Function

Public Function GetLastByteArray(ByRef destArray() As Byte, ByVal arrayIndex As Long, ByVal count As Long)
    Dim realCount&, i&
    If count < size_ Then
        realCount = count
    Else
        realCount = size_
    End If
    For i = (realCount - 1) To 0 Step -1
        destArray((arrayIndex + i) - 1) = GetLastByte()
    Next
    GetLastByteArray = realCount
End Function

Public Function PeekByte() As Byte
    Dim item As Byte
    If IsEmpty = True Then
        Debug.Print "The buffer is empty."
        Exit Function
    End If
    item = buffer_(head_)
    PeekByte = item
End Function
Public Function PeekByteArray(ByVal count As Long) As Byte()
    Dim items() As Byte
    ReDim items(count)
    If IsEmpty = True Then
        Debug.Print "The buffer is empty."
        Exit Function
    End If
    CopyTo1 items
    PeekByteArray = items
End Function
Public Function PeekAtByte(ByVal index As Long) As Byte
    If IsEmpty = True Then
        Debug.Print "The buffer is empty."
        Exit Function
    End If
    If index < 0 Or index >= size_ Then
        Debug.Print "Index must be between 0 and " & CStr(size_) & "."
        Exit Function
    End If
    PeekAtByte = buffer_(GetHeadIndex(index))
End Function
Public Function PeekLastByte() As Byte
    Dim item As Byte
    Dim index&
    If IsEmpty = True Then
        Debug.Print "The buffer is empty."
        Exit Function
    End If
    index = GetTailIndex(0)
    item = buffer_(index)
    PeekLastByte = item
End Function
Public Function PeekLastByteArray(ByRef destArray() As Byte, ByVal arrayIndex As Long, ByVal count As Long) As Long
    Dim realCount&, i&
    If count < size_ Then
        realCount = count
    Else
        realCount = size_
    End If
    For i = 0 To (realCount - 1)
        destArray(arrayIndex + (realCount - (i + 1))) = buffer_(GetTailIndex(i))
    Next
    PeekLastByteArray = realCount
End Function
Public Function PutByteArray(ByRef srcArray() As Byte, ByVal arrayIndex As Long, ByVal count As Long)
    If allowOverwrite_ = False And count > (capacity_ - size_) Then
        Debug.Print "The buffer does not have sufficient capacity to put new items."
        Exit Function
    End If
    Dim i&
    For i = 0 To (count - 1)
        PutByte srcArray(arrayIndex + i)
    Next
    PutByteArray = i
End Function
Public Sub PutByte(item As Byte)
    If allowOverwrite_ = False And size_ = capacity_ Then
        Debug.Print "The buffer does not have sufficient capacity to put new items."
        Exit Sub
    End If
    buffer_(tail_) = item
    tail_ = tail_ + 1
    If size_ = capacity_ Then
        head_ = head_ + 1
        If head_ >= capacity_ Then
            head_ = head_ - capacity_
        End If
    End If
    If tail_ = capacity_ Then
        tail_ = 0
    End If
    If size_ <> capacity_ Then
        size_ = size_ + 1
    End If
End Sub
Public Sub Skip(ByVal count As Long)
    head_ = GetHeadIndex(count)
End Sub
Public Function ToArray() As Byte()
    Dim result() As Byte
    ReDim result(size_)
    CopyTo1 result
End Function
Private Function GetHeadIndex(ByVal index As Long) As Long
    Dim newIndex&
    newIndex = head_ + index
    If (newIndex >= capacity_) Then
        newIndex = newIndex - capacity_
    End If
    GetHeadIndex = newIndex
End Function
Private Function GetTailIndex(ByVal index As Long) As Long
    Dim bufferIndex&
    If tail_ = 0 Then
        bufferIndex = size_ - (index + 1)
    Else
        bufferIndex = tail_ - (index + 1)
    End If
    GetTailIndex = bufferIndex
End Function

Private Sub Class_Initialize()
    allowOverwrite_ = False
    capacity_ = 0
    head_ = 0
    size_ = 0
    tail_ = 0
End Sub
