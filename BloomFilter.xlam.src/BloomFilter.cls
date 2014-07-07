VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BloomFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

''
' BloomFilter
' (c) Rhys Kidd - https://github.com/Echelon9/bloomfilter.vba
'
' Implementation of the BloomFilter class
'
' @author: rhyskidd@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Option Explicit

' Declare private members
Private pBuckets() As Long
Private p_Locations() As Long
Private pK As Integer
Private pM As Integer

' Class initializer -- VBA cannot accept parameters to this Sub
Private Sub Class_Initialize()
    
    Dim m, k As Integer
    
    '==== Default constructor values
    m = 32 * 256    ' Number of bits. Must be a multiple of 32.
    k = 4           ' Number of hashing functions.
    '====
    
    Dim n, i As Integer
 
    ' Error checking to ensure m is a multiple of 32.
    n = Ceiling(m / 32)
    m = n * 32
 
    ' Assign the class global values
    pM = m
    pK = k
    
    ReDim pBuckets(0 To m) As Long
    ReDim p_Locations(0 To k) As Long
    
End Sub

' Class destructor -- VBA cannot accept parameters to this Sub
Private Sub Class_Terminate()

    Erase pBuckets, p_Locations

End Sub

' Add strItem to BloomFilter
Public Function Add(ByVal strItem As String)
    
    Dim l As Variant
    Dim i, k As Integer
    
    l = Locations(strItem & "")
    i = 0
    k = pK
    
    While i < k
        ' buckets[Math.floor(l[i] / 32)] |= 1 << (l[i] % 32);
        pBuckets(Floor(l(i) / 32)) = pBuckets(Floor(l(i) / 32)) Xor shl(1, (l(i) Mod 32))
        
        i = i + 1
    Wend
    
End Function

' Test strItem in BloomFilter
Public Function Test(ByVal strItem As String) As Boolean
    
    Dim l As Variant
    Dim i, k As Integer
    Dim b As Long
    
    l = Locations(strItem & "")
    i = 0
    k = pK
    
    While i < k
        b = l(i)
        
        ' (buckets[Math.floor(b / 32)] & (1 << (b % 32))) === 0
        If ((pBuckets(Floor(b / 32)) And shl(1, (b Mod 32))) = 0) Then
            Test = False
            GoTo Test_Exit
        End If
        
        i = i + 1
    Wend
    
    Test = True

Test_Exit:
    ' Fast path exit
    Exit Function
    
End Function

' Estimated cardinality.
Public Property Get Size() As Double
    
    On Error GoTo Size_Err
    
    Dim bits As Variant
    Dim i As Integer
    
    bits = 0
    
    ' for (var i = 0, n = buckets.length; i < n; ++i) bits += popcnt(buckets[i]);
    For i = LBound(pBuckets) To UBound(pBuckets)
        bits = bits + popcnt(pBuckets(i))
    Next i
    
    Size = -pM * Log(1 - bits / pM) / pK
    
    Exit Property
    
Size_Err:
    MsgBox "Warning: BloomFilter.Size() not fully implemented"
    Size = -1
    
End Property

' See http://willwhim.wordpress.com/2011/09/03/producing-n-hash-functions-by-hashing-only-once/
' and http://citeseerx.ist.psu.edu/viewdoc/summary?doi=10.1.1.152.579
'
'    Less hashing, same performance: Building a better bloom filter (2006)
'    Adam Kirsch, Michael Mitzenmacher
'
'    "A standard technique from the hashing literature is to use two hash functions h1(x)
'    and h2(x) to simulate additional hash functions of the form gi(x) = h1(x) + ih2(x).
'    We demonstrate that this technique can be usefully applied to Bloom filters and
'    related data structures. Specifically, only two hash functions are necessary to
'    effectively implement a Bloom filter without any loss in the asymptotic false
'    positive  probability. This leads to less computation and potentially less need
'    for randomness in practice."
'
Private Property Get Locations(ByVal v As String) As Variant
        
    Dim k, m, a, b, i, X As Integer
        
    k = pK
    m = pM
    a = FNV1a(v)
    b = FNV1a(a)
    i = 0
    X = a Mod m
    
    While i < k
        p_Locations(i) = IIf(X < 0, (X + m), X)
        X = (X + b) Mod m
        
        i = i + 1
    Wend
    
    Locations = p_Locations
    
End Property

' http://graphics.stanford.edu/~seander/bithacks.html#CountBitsSetParallel
Private Function popcnt(ByVal v As Long) As Variant
    ' Note: There is no Subtraction assignment (-=) operator in VBA by default
    
    v = v - (shr(v, 1) And &H55555555)
    v = (v And &H33333333) + (shr(v, 2) And &H33333333)
    v = shr((v + shr(v, 4) And &HF0F0F0F) * (&H1010101), 24)
    
    popcnt = v
End Function

' http://www.excely.com/excel-vba/bit-shifting-function.shtml
Private Function shr(ByVal Value As Long, ByVal Shift As Byte) As Long
    Dim i As Byte
    shr = Value
    If Shift > 0 Then
        shr = Int(shr / (2 ^ Shift))
    End If
End Function

' http://www.excely.com/excel-vba/bit-shifting-function.shtml
Private Function shl(ByVal Value As Long, ByVal Shift As Byte) As Long
    shl = Value
    If Shift > 0 Then
        Dim i As Byte
        Dim m As Long
        For i = 1 To Shift
            m = shl And &H40000000
            shl = (shl And &H3FFFFFFF) * 2
            If m <> 0 Then
                shl = shl Or &H80000000
            End If
        Next i
    End If
End Function

' Fowler-Noll-Vo (FNV-1a) hash implementation.
'
' FNV is a non-cryptographic hash function created by Glenn Fowler, Landon Curt Noll, and Phong Vo.
'
' The FNV hash algorithms and sample FNV source code have been released into the public domain.
' http://www.isthe.com/chongo/tech/comp/fnv/index.html#public_domain
'
Private Function FNV1a(ByVal source As String) As Long
  Dim Map() As Byte
  Dim i As Long
  Dim HashHi As Long
  Dim HashLo As Long
  Dim HiBit As Long
  Dim Max As Long
  
  Const PHI As Long = &H100
  Const PLO As Long = &H193

  If LenB(source) = 0 Then Exit Function
  Map = StrConv(source, vbFromUnicode)
  HashLo = &H9DC5&
  HashHi = &H811C&
  Max = UBound(Map)
  
  Do
    HashLo = (HashLo And &HFFFF&) Xor Map(i)
    HashHi = ((HashHi * PLO) + (HashLo * PHI)) And &HFFFF&
    HashLo = (HashLo * PLO)
    HashHi = HashHi + (HashLo \ &H10000)
    i = i + 1
  Loop While i <= Max
  
  HiBit = HashHi And &H8000&
  HashHi = ((HashHi And &H7FFF&) * &H10000) Or (HashLo And &HFFFF&)
  If HiBit Then HashHi = HashHi Or &H80000000

  FNV1a = HashHi
End Function

Private Function Ceiling(ByVal X As Double, Optional ByVal Factor As Double = 1) As Double
    ' X is the value you want to round
    ' is the multiple to which you want to round
    Ceiling = (Int(X / Factor) - (X / Factor - Int(X / Factor) > 0)) * Factor
End Function

Private Function Floor(ByVal X As Double, Optional ByVal Factor As Double = 1) As Double
    ' X is the value you want to round
    ' is the multiple to which you want to round
    Floor = Int(X / Factor) * Factor
End Function

