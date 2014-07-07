Bloom Filter
============

This Visual Basic for Applications bloom filter implementation uses the 
non-cryptographic [Fowler–Noll–Vo hash function][1] for speed.

Usage
-----

    Dim bloom As New BloomFilter
    
    Set bloom = New BloomFilter

    ' Add some elements to the filter.
    bloom.Add ("foo")
    bloom.Add ("bar")

    ' Test if an item is in our filter.
    ' Returns true if an item is probably in the set,
    ' or false if an item is definitely not in the set.
    bloom.Test ("foo")
    bloom.Test ("bar")
    bloom.Test ("blah")


Implementation
--------------

Although the bloom filter requires *k* hash functions, we can simulate this
using only *two* hash functions.  In fact, we cheat and get the second hash
function almost for free by iterating once more on the first hash using the FNV
hash algorithm.

Thanks to Will Fitzgerald for his [help and inspiration][2] with the hashing
optimisation.

[1]: http://isthe.com/chongo/tech/comp/fnv/
[2]: http://willwhim.wordpress.com/2011/09/03/producing-n-hash-functions-by-hashing-only-once/
