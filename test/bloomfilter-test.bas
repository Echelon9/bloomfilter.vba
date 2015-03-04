
''
' BloomFilter-Test
' (c) Rhys Kidd - https://github.com/Echelon9/bloomfilter.vba
'
' General specs for the BloomFilter class
'
' @author: rhyskidd@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "BloomFilter"
    
    Dim f As New BloomFilter

    With Specs.It("basic bloom filter")
    
        Set f = New BloomFilter

        Dim n1 As String: n1 = "Bess"
        Dim n2 As String: n2 = "Jane"

        f.Add (n1)

        .Expect(f.Test(n1)).ToEqual True
        .Expect(f.Test(n2)).ToEqual False

    End With

    With Specs.It("jabberwocky")
    
        Set f = Nothing
        Set f = New BloomFilter

        Dim jabberwocky As String: jabberwocky = "`Twas brillig, and the slithy toves\n  Did gyre and gimble in the wabe:\nAll mimsy were the borogoves,\n  And the mome raths outgrabe.\n\n\'Beware the Jabberwock, my son!\n  The jaws that bite, the claws that catch!\nBeware the Jubjub bird, and shun\n  The frumious Bandersnatch!\'\n\nHe took his vorpal sword in hand:\n  Long time the manxome foe he sought --\nSo rested he by the Tumtum tree,\n  And stood awhile in thought.\n\nAnd, as in uffish thought he stood,\n  The Jabberwock, with eyes of flame,\nCame whiffling through the tulgey wood,\n  And burbled as it came!\n\nOne, two! One, two! And through and through\n  The vorpal blade went snicker-snack!\nHe left it dead, and with its head\n  He went galumphing back.\n\n\'And, has thou slain the Jabberwock?\n  Come to my arms, my beamish boy!\nO frabjous day! Callooh! Callay!'\n  He chortled in his joy.\n\n`Twas brillig, and the slithy toves\n  Did gyre and gimble in the wabe;"
        
        n1 = jabberwocky
        n2 = jabberwocky & "\n"

        f.Add (n1)

        .Expect(f.Test(n1)).ToEqual True
        .Expect(f.Test(n2)).ToEqual False

    End With

    With Specs.It("basic uint32")
    
        Set f = Nothing
        Set f = New BloomFilter
        Dim n3 As String

        ' Will need to confirm the data type in VBA
        n1 = "\u0100"
        n2 = "\u0101"
        n3 = "\u0103"

        f.Add (n1)

        .Expect(f.Test(n1)).ToEqual True
        .Expect(f.Test(n2)).ToEqual False
        .Expect(f.Test(n3)).ToEqual False

    End With

    With Specs.It("wtf")
    
        Set f = Nothing
        Set f = New BloomFilter

        f.Add ("abc")

        .Expect(f.Test("wtf")).ToEqual False

    End With

    With Specs.It("works with integer types")
    
        Set f = Nothing
        Set f = New BloomFilter

        f.Add (1)

        .Expect(f.Test(1)).ToEqual True
        .Expect(f.Test(2)).ToEqual False

    End With

    With Specs.It("size")

        Set f = Nothing
        Set f = New BloomFilter
        Dim i As Integer: i = 0

        For i = 0 To 100
            f.Add (i)
        Next i
        .Expect(f.Size()).ToBeCloseTo 99.61766, 5

        For i = 100 To 1000
            f.Add (i)
        Next i
        .Expect(f.Size()).ToBeCloseTo 943.51438, 5

    End With

    Set f = Nothing
    
    ' InlineRunner.RunSuite Specs
    ' InlineRunner.RunSuite Specs, ShowFailureDetails:=True, ShowPassed:=True, ShowSuiteDetails:=True
End Function
