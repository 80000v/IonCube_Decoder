Attribute VB_Name = "Module1"
Option Explicit
Option Base 0
'The above "Option X" commands must be the 1st.lines in the file.
'Option Compare Database    'DO NOT USE IN THIS MODULE, not even for MS Access



'This is the Visual Basic for Applications (VBA) version of the  MT19937ar,
'or   "MERSENNE TWISTER"   algorithm for pseudo random number generation,
'with initialization improved, by  MAKOTO MATSUMOTO  and  TAKUJI NISHIMURA,
'of  2002/1/26.
'This translation to VBA was made and tested by Pablo Mariano Ronchi (2005-Sep-12)

'Note 1: VBA is the Visual Basic language used in MS Access, MS Excel and, in general,
'        in MS Office, and is called simply "Visual Basic" or VBA, hereinafter.
'Note 2: This same code compiles in Visual Basic (VB) without modifications.



'Please read the comments about this VBA version that follow the ones below, by
'the authors of the "MERSENNE TWISTER" algorithm.






'/*
'   A C-program for MT19937, with initialization improved 2002/1/26.
'   Coded by Takuji Nishimura and Makoto Matsumoto.
'
'   Before using, initialize the state by using init_genrand(seed)
'   or init_by_array(init_key, key_length).
'
'   Copyright (C) 1997 - 2002, Makoto Matsumoto and Takuji Nishimura,
'   All rights reserved.
'
'   Redistribution and use in source and binary forms, with or without
'   modification, are permitted provided that the following conditions
'   are met:
'
'     1. Redistributions of source code must retain the above copyright
'        notice, this list of conditions and the following disclaimer.
'
'     2. Redistributions in binary form must reproduce the above copyright
'        notice, this list of conditions and the following disclaimer in the
'        documentation and/or other materials provided with the distribution.
'
'     3. The names of its contributors may not be used to endorse or promote
'        products derived from this software without specific prior written
'        permission.
'
'   THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
'   "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
'   LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
'   A PARTICULAR PURPOSE ARE DISCLAIMED.  IN NO EVENT SHALL THE COPYRIGHT OWNER OR
'   CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL,
'   EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
'   PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR
'   PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF
'   LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING
'   NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
'   SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'
'
'   Any feedback is very welcome.
'   http://www.math.sci.hiroshima-u.ac.jp/~m-mat/MT/emt.html
'   email: m-mat @ math.sci.hiroshima-u.ac.jp (remove space)
'*/



'COMMENTS ABOUT THIS VISUAL BASIC FOR APPLICATIONS (VBA) VERSION OF
'THE  "Mersenne Twister"  ALGORITHM:
'
'- All the statements made by the authors of the original algorithm and implementation,
'  present in the original C code and copied above, including but not limiting to those
'  regarding the license, the usage without any warranties, and the conditions for
'  distribution, apply to this Visual Basic program for testing that algorithm.
'
'- If you use this Visual Basic version, just for the record please send me an email to:
'
'  pmronchi@yahoo.com.ar
'
'  with an appropriate reference to your work, the city and country where you reside,
'  and the word "MT19937ar" in the subject. Thanks.
'
'
'- FUNCTIONS AND PROCEDURES IMPLEMENTED:
'
'       Function                    Returns values in the range:
'       -------------------------   -------------------------------------------------------
'       genrand_int32()             [0, 4294967295]  (0 to 2^32-1, A Double OF INTEGER VALUE)
'
'       genrand_int31()             [0, 2147483647]  (0 to 2^31-1)
'
'       genrand_real1()             [0.0, 1.0]   (both 0.0 and 1.0 included)
'
'       genrand_real2()             [0.0, 1.0) = [0.0, 0.9999999997672...]
'                                                (0.0 included, 1.0 excluded)
'
'       genrand_real3()             (0.0, 1.0) = [0.0000000001164..., 0.9999999998836...]
'                                                (both 0.0 and 1.0 excluded)
'
'       genrand_res53()             [0.0,~1.0] = [0.0, 1.00000000721774...]
'                                                (0.0 included, ~1.0 included)
'
'       The following ADDITIONAL functions
'       ARE NOT PRESENT IN THE ORIGINAL C CODE:
'
'       NOTE: the limits shown below, marked with (*), are valid if gap==5.0e-13
'
'
'       genrand_int32SignedLong()  [-2147483648, 2147483647]   (-2^31 to 2^31-1)
'
'       genrand_real2b()           [0.0, 1.0)=[0, 1-(2*gap)] =[0.0, 0.9999999999990] (*)
'                                             (0.0 included, 1.0 excluded)
'
'       genrand_real2c()           (0.0, 1.0]=[0+(2*gap),1.0]=[1.0e-12, 1.0] (*)
'                                             (0.0 excluded, 1.0 included)
'
'       genrand_real3b()           (0.0, 1.0)=[0+gap, 1-gap] =[5.0e-13, 0.9999999999995] (*)
'                                             (both 0.0 and 1.0 excluded)
'
'
'       (See the "Acknowledgements" section for the following functions)
'
'       genrand_real4b()           [-1.0,1.0]=[-1.0, 1.0]
'                                             (-1.0 included, 1.0 included)
'
'       genrand_real5b()           (-1.0,1.0)=[-1.0+(2*gap), 1.0-(2*gap)]=
'                                                    [-0.9999999999990,0.9999999999990] (*)
'                                             (-1.0 excluded, 1.0 excluded)
'
'
'       Procedure                 Arguments
'       ------------------------  ---------------------------------------------------------
'       init_genrand(seed)        any seed included in [-2147483648, 2147483647]
'       init_by_array(array,len)  array has N elements of type Long;
'                                 len==N, is the number of elements in the array
'
'
'- USAGE:
'
'  In your Visual Basic application (MS Access, MS Excel, VB compiler, etc):
'
'  1) Add this source file as a module with the proper name (I suggest "mt19937ar").
'     If you want to use the Mersenne Twister algorithm, and not just to test it,
'     REMOVE THE main() FUNCTION, at the end of this file, and keep the rest untouched.
'
'  2) (Optionally) call either the procedure init_genrand() with a proper seed,
'     or init_by_array() with the proper arguments, in order to (optionally)
'     initialize the pseudorandom sequence.
'
'     Note: If init_genrand() or init_by_array() were not called previous to the call
'           to any of the genrand_X() functions listed above, anyone of them will
'           automatically call init_genrand() by default the first time that is used,
'           with a seed==kDefaultSeed==5489 . This feature is present in the original C code.
'
'  3) Then, simply call any of the genrand_X() functions listed above.
'

'  Example of usage:
'
'     'Optionally initialize the pseudorandom sequence in the following way:
'
'          init_genrand 69769        'any seed in the range [-2147483648, 2147483647]
'
'     'Or (optionally, too) declare an array, initialize it, and call init_by_array()
'     'with the array and the number of its elements as arguments, in the following way:
'
'          Dim init As Variant: init = Array(123, 234, 345, 456)
'          Dim length As Long: length = 4
'
'          init_by_array init, length
'
'     '(If you do not use any of the 2 procedures above to initialize, initialization
'     'is performed automatically by a call to: init_genrand 5489)
'
'
'     'Then, get a pseudorandom number and assign it to variable x (of the proper
'     'Long or Double type), by calling one of the genrand_X() functions:
'
'          Dim x As Double
'          x=genrand_real2()
'
'
'- ON THE NEED OF THE Mersenne Twister ALGORITHM:
'
'  If you are serious about (pseudo) randomness in your Visual Basic project, then
'  you MUST use this Visual Basic version of the Mersenne Twister algorithm.
'  DO NOT RELY on the randomize() and rnd() routines provided by Visual Basic!
'
'
'- ON PERFORMANCE:
'
'  - If you are using Visual Basic, then performance is hardly an issue, and very probable
'    these routines will not be a bottleneck. Anyway, I made my best to optimize the code
'    for speed without compromising its clarity and readability, and while strictly
'    following the original C code.
'    If you need better performance, consider using the C or PHP version of the
'    Mersenne Twister algorithm, and MySQL for database management, for instance.
'
'  - This new version (September 2005), is around 9% faster than the previous one
'    published on-line (April 2005).
'
'  - Performance tests: For my own testing procedures I have developed a full test.
'    I would like to make the test package publicly available, but as of this writing
'    I have not asked Mr.Makoto Matsumoto yet. If he agrees, you will probably
'    find it close to the place were you found this code.
'
'  - Just to give you an idea of the difference in performance between C and VBA, one
'    of the parts of my test is a loop that only calls genrand_int32() repeatedly,
'    generating 100 million (1e08) pseudorandom numbers.
'    The times in my old Pentium MMX, 120Mhz, for this part of the test are:
'         6008 seconds (1h40m) for VBA, and 56 seconds for C (a relation of 107:1).
'
'  - A simple VBA code to test the performance is given below (SimplePerformanceTest()).
'    The C code provided in the main() function by the authors of the Mersenne Twister
'    algorithm could be easily adapted to match the result of a suitable modified
'    SimplePerformanceTest():
'
'    Public Sub SimplePerformanceTest()
'    Const kMaxNr As Long = 1000000    'CHANGE THIS VALUE AS NEEDED
'    Dim ii As Long, tmp As Double, sec1 As Double, sec2 As Double
'    Open "mt19937arVBtest.txt" For Output As #1    'open the output file
'
'    sec1 = Timer
'    For ii = 1 To kMaxNr: tmp = genrand_int32(): Next 'use any of the functions
'    sec2 = Timer: tmp = sec2 - sec1
'
'    Print #1, "Elapsed time in seconds for generating "; Trim(kMaxNr); " numbers: "; _
'              Fix(tmp) & "." & Format(Fix((tmp - Fix(tmp) + 0.005) * 100), "00")
'    Close #1    'close the output file
'    End Sub 'SimplePerformanceTest
'
'
'- DIFFERENCES WITH THE ORIGINAL C FUNCTIONS AND SOURCE FILE:
'
'  - The VBA version of genrand_int32() returns a Double in the range [0, 4294967295],
'    instead of a Long, because the Long type cannot accomodate the range of unsigned long
'    values returned by the C version. But all the returned values of the VBA version are
'    integers, as it is due. (To be true, they are not, but they should be considered so
'    because they are as integer as an integer value can be represented in a Double variable
'    following the IEEE standard, which is as exact as we need for all practical uses).
'    Because of this change, I had to rename the original C function genrand_int32() as
'    genrand_int32SignedLong(), returning a Long in the range [-2147483648,2147483647].
'
'  - I added uAdd(), uMult() and uDiv() to emulate unsigned long addition, multiplication
'    and division WITHIN THE CONTEXT of the Mersenne Twister algorithm implementation in
'    Visual Basic. THEY ARE NOT INTENDED FOR GENERAL USE!.
'
'  - There are minor changes (addition and use of variable mtb) to produce the same result
'    in Visual Basic as in C, when a genrand_X() function is called without a previous
'    call to one of the init_X procedures. I apologize for using this not very elegant
'    patch, but there is no simpler way to simulate the use of the "static" word in C,
'    given that the VBA's "static" word does not behave in a similar way.
'    Another minor change, for similar reasons, is the declaration and initialization of
'    the mag01() array.
'
'  - I added many constants, for clarity in some cases, and also for speed in others,
'    because some operations could be faster if defined as a multiplication instead of
'    a division, in some processors.
'
'  - I made small changes in the main() function, in order to easily change the number
'    of output values in the listings.
'
'
'- IS THIS VBA CODE DEPENDABLE?
'
'  Well, I think so. I have tested the code by generating 101026001 pseudorandom numbers
'  using all of the functions, printed 389 of them, and compared the result with a similar
'  test I wrote in C using the original code of the Mersenne Twister algorithm. Both
'  outputs were exactly the same, except for the timings and some variable headings.
'
'  Besides, I kept the original C code commented out in this source file, in order
'  to easily check and understand the translation to Visual Basic. So, you will find
'  a block of one or a few lines of Visual Basic code following the corresponding
'  -commented out- original C block.
'  This way you can easily see that I have strictly followed the original C code, except
'  for the minor differences explained in the above section, and verify that the VBA code
'  is an almost exact "copy" of the original algorithm.
'  This fact provides another level of confidence to the end result.
'
'  No bugs or problems were reported since the publication on-line of the previous
'  version (April 2005).
'
'
'- ACKNOWLEDGEMENTS:
'
'  I want to thank Mr.MAKOTO MATSUMOTO and Mr.TAKUJI NISHIMURA for creating and
'  generously sharing their excellent algorithm.
'
'  I also want to thank my friends Alejandra María Ribichich, Mariana Francisco Mera,
'  and Víctor Fernando Torres, for their inspiration and support.
'
'  My friend Claudio Pacciarini clarified me some of the differences between VB and VBA.
'
'  Mr.Mutsuo Saito, assistant to professor Matsumoto, patiently e-mailed with me and
'  performed the necessary tasks for the previous version (April 2005) to appear on-line.
'
'  Mr. Kenneth C. Ives (USA) and Mr. David Grundy (Hong Kong) were the first ones
'  aknowledging the use of this VBA code. Thanks for your feedback!
'
'  Mr. Kenneth C. Ives also sent me some code and the idea in which I based
'  genrand_real4b() and genrand_real5b()
'
'
'  Pablo Mariano Ronchi
'  Buenos Aires, Argentina
'
'
'
'End of comments for the Visual Basic version





'#include <stdio.h>
'
'/* Period parameters */
'#define N 624
'#define M 397
'#define MATRIX_A 0x9908b0dfUL   /* constant vector a */
'#define UPPER_MASK 0x80000000UL /* most significant w-r bits */
'#define LOWER_MASK 0x7fffffffUL /* least significant r bits */
'
'static unsigned long mt[N];     /* the array for the state vector */
'static int mti=N+1;             /* mti==N+1 means mt[N] is not initialized */
Const N As Long = 624
Const M As Long = 397
Const MATRIX_A As Long = &H9908B0DF     '/* constant vector a */
Const UPPER_MASK As Long = &H80000000   '/* most significant w-r bits */
Const LOWER_MASK As Long = &H7FFFFFFF   '/* least significant r bits */

'To avoid innecesary operations while using the Visual Basic interpreter:
Const kDiffMN As Long = M - N
Const Nuplim As Long = N - 1
Const Muplim As Long = M - 1
Const Nplus1 As Long = N + 1
Const NuplimLess1 As Long = Nuplim - 1
Const NuplimLessM As Long = Nuplim - M

'static unsigned long mt[N];  /* the array for the state vector */
'static int mti=N+1;          /* mti==N+1 means mt[N] is not initialized */
Dim mt(0 To Nuplim) As Long  '/* the array for the state vector */
Dim mti As Long

'In the C original version the following array, mag01(), is declared within
'the function genrand_int32(). In VBA I had to declare it global for performance
'considerations, and because there is no way in VBA to emulate the use of the word
'"static" in C:
'
'static unsigned long mag01[2]={0x0UL, MATRIX_A};
'/* mag01[x] = x * MATRIX_A  for x=0,1 */
Dim mag01(2) As Long

Dim mtb As Boolean   'needed in Visual Basic

'Other constants defined to be used in this Visual Basic version:

'Powers of 2: k2_X means 2^X
Const k2_8 As Long = 256
Const k2_16 As Long = 65536
Const k2_24 As Long = 16777216

Const k2_31 As Double = 2147483648#     '2^31   ==  2147483648 == 80000000
Const k2_31Neg As Double = 0# - k2_31   '-2^31  == -2147483648 == 80000000
Const k2_31b As Double = k2_31 - 1#     '2^31-1 ==  2147483647 == 7FFFFFFF
Const k2_32 As Double = 2# * k2_31      '2^32   ==  4294967296 == 0
Const k2_32b As Double = k2_32 - 1#     '2^32-1 ==  4294967295 == FFFFFFFF == -1


'Constants for shift left operation:
Const kShl7 As Long = 128          '128==2^7
Const kShl15 As Long = 32768       '32768==2^15

'Constants for shift right operation:
Const kShr1 As Long = 2            '2==2^1
Const kShr5 As Long = 32           '32==2^5
Const kShr6 As Long = 64           '64==2^6
Const kShr11 As Long = 2048        '2048==2^11
Const kShr18 As Long = 262144      '262144==2^18
Const kShr30 As Long = 1073741824  '1073741824==2^30  used in init_X() functions

'The following constant has its value defined by the authors of the
'Mersenne Twister algorithm
Const kDefaultSeed As Long = 5489


'The following constant, is used within genrand_real1(), which returns values in [0,1]
Const kMT_1 As Double = 1# / k2_32b

'The following constant, is used within genrand_real2(), which returns values in [0,1)
Const kMT_2 As Double = 1# / k2_32

'The following constant, is used within genrand_real3(), which returns values in (0,1)
Const kMT_3 As Double = kMT_2

'The following constant, used within genrand_res53(), is needed, because the Visual
'Basic interpreter cannot read real LITERALS with the the same precision as a C compiler,
'and so ends up truncating the least significant decimal digit(s), a '2' in this case.
'The original factor used in the C code is: 9007199254740992.0
Const kMT_res53 As Double = 1# / (9.00719925474099E+15 + 2#)    'add lost digit '2'


'The following constants, are used within the ADDITIONAL functions genrand_real2b() and
'genrand_real3b(), equivalent to genrand_real() and genrand_real3(), but that return
'evenly distributed values in the ranges [0, 1-kMT_Gap] and [0+kMT_Gap, 1-kMT_Gap],
'respectively. A similar statement is valid also for genrand_real2c(), genrand_real4b()
'and genrand_real5b(). See the section "Functions and procedures implemented" above,
'for more details.
'
'If you want to change the value of kMT_Gap, it is suggested to do it so that:
'   5e-15 <= kMT_Gap <= 5e-2

Const kMT_Gap As Double = 0.0000000000005       '5.0E-13
Const kMT_Gap2 As Double = 2# * kMT_Gap         '1.0E-12
Const kMT_GapInterval As Double = 1# - kMT_Gap2 '0.9999999999990

Const kMT_2b As Double = kMT_GapInterval / k2_32b
Const kMT_2c As Double = kMT_2b
Const kMT_3b As Double = kMT_2b
Const kMT_4b As Double = 2# / k2_32b
Const kMT_5b As Double = (2# * kMT_GapInterval) / k2_32b   '1.999999999998/k2_32b



'Just for source file formatting. To make a space between the line separation below
'and the above declarations:
Const EndOfConstVarSection As Byte = 0








Private Function uAdd(ByVal x As Long, ByVal y As Long) As Long
'Unsigned Add: adds the two (signed) Long parameters, treated as
'unsigned long, and returns the result as a (signed) Long result:

Dim tmp As Double

tmp = CDbl(x) + y

If tmp < k2_31Neg Then
    uAdd = CLng(k2_32 + tmp)
Else
    If tmp > k2_31b Then
        uAdd = CLng(tmp - k2_32)
    Else
        uAdd = CLng(tmp)
    End If
End If

End Function    'uAdd





Private Function uMult(ByVal x As Long, ByVal y As Long) As Long
'Unsigned Mult: multiplies the two (signed) Long parameters, treated as
'unsigned long, and returns the lowest 4 bytes of the 8 bytes result
'as a (signed) Long result:

'This function emulates the multiplication of two (unsigned long) numbers,
'and is needed because the type Double has only a 53 bits mantissa
'and the result of multiplying 2 Long variables might need 64 bits to accurately
'represent the result in some cases.

'x   == ABCD  == A000 + B00 + C0 + D
'y   == EFGH  == E000 + F00 + G0 + H

'       Note: in the following, "ae" means the 2 bytes result of A*E, "bg" of B*G, etc:
'
'x*y == ae 000 000 +  'discard, result is too high
'       af 000  00 +  'discard, result is too high
'       ag 000   0 +  'discard, result is too high
'       ah 000     +  'take lowest byte

'       be  00 000 +  'discard, result is too high
'       bf  00  00 +  'discard, result is too high
'       bg  00   0 +  'take lowest byte
'       bh  00     +  'take both bytes

'       ce   0 000 +  'discard, result is too high
'       cf   0  00 +  'take lowest byte
'       cg   0   0 +  'take both bytes
'       ch   0     +  'take both bytes

'       de     000 +  'take lowest byte
'       df      00 +  'take both bytes
'       dg       0 +  'take both bytes
'       dh            'take both bytes


'Given that "aa" and "ee" are used just once, they are replaced by their
'values: aa == (x \ k2_24) and ee == (y \ k2_24), and they are not declared:
'Dim aa As Long, ee As Long

Dim bb As Long, cc As Long, dd As Long
Dim ff As Long, gg As Long, hh As Long
Dim r3 As Long, r2 As Long, r1 As Long, r0 As Long
Dim tmp As Double


'x==ABCD, y==EFGH
bb = (x \ k2_16) Mod k2_8: cc = (x \ k2_8) Mod k2_8: dd = x Mod k2_8
ff = (y \ k2_16) Mod k2_8: gg = (y \ k2_8) Mod k2_8: hh = y Mod k2_8


'get the 1st (lowest) byte of the result, r0:
'       dh             'take both bytes
r0 = dd * hh

'get the 2nd byte of the result, r1, and add carry from r0:
'       ch   0      +  'take both bytes
'       dg        0    'take both bytes
r1 = cc * hh + dd * gg + r0 \ k2_8

'get the 3rd byte of the result, r2, and add carry from r1:
'       bh  00      +  'take both bytes
'       cg   0    0 +  'take both bytes
'       df       00    'take both bytes
r2 = bb * hh + cc * gg + dd * ff + r1 \ k2_8

'get the 4th (highest) byte of the result, r3, and add carry from r2:
'       ah 000      +  'take lowest byte
'       bg  00    0 +  'take lowest byte
'       cf   0   00 +  'take lowest byte
'       de      000    'take lowest byte
r3 = (((x \ k2_24) * hh + bb * gg + cc * ff + dd * (y \ k2_24)) Mod k2_8) + r2 \ k2_8


'tmp = CDbl(r3) * k2_24 + r2 * k2_16 + r1 * k2_8 + r0
tmp = CDbl(r3 Mod k2_8) * k2_24 + (r2 Mod k2_8) * k2_16 + (r1 Mod k2_8) * k2_8 + (r0 Mod k2_8)

'now we have a 32 bits number (tmp) that can be processed without losing precision
'using the 53 bits mantissa of the Double type

If tmp < k2_31Neg Then
    uMult = CLng(k2_32 + tmp)
Else
    If tmp > k2_31b Then
        uMult = CLng(tmp - k2_32)
    Else
        uMult = CLng(tmp)
    End If
End If

End Function    'uMult




Private Function uDiv(ByVal x As Long, ByVal y As Long) As Long
'Unsigned Divide: divides the two (signed) Long parameters, treated as
'unsigned long, and returns the result as a (signed) Long result:

'No need to check y: this function is always called with y>=2.0
'If y < 0 Then y = k2_32 + y End If

If x < 0 Then
    uDiv = CLng(Fix((k2_32 + x) / y))
Else
    uDiv = CLng(Fix(x / y))
End If

End Function    'uDiv




Private Function uDiv2(ByVal x As Double, ByVal y As Long) As Double
'Unsigned Divide, 2nd.definition: divides a Double x by a (signed) Long divisor y,
'treated as unsigned long, and returns the result as a Double of integer value:

'No need to check y: this function is always called with y>=2.0
'If y < 0 Then y = k2_32 + y End If

If x < 0 Then
    uDiv2 = Fix((k2_32 + x) / y)
Else
    uDiv2 = Fix(x / y)
End If

End Function    'uDiv2








Public Sub init_genrand(ByVal seed As Long)      'void init_genrand(unsigned long s)
'/* initializes mt[N] with a seed */
'mt[0]= s & 0xffffffffUL;
'for (mti=1; mti<N; mti++) {
'    mt[mti] =
'    (1812433253UL * (mt[mti-1] ^ (mt[mti-1] >> 30)) + mti);
'    /* See Knuth TAOCP Vol2. 3rd Ed. P.106 for multiplier. */
'    /* In the previous versions, MSBs of the seed affect   */
'    /* only MSBs of the array mt[].                        */
'    /* 2002/01/09 modified by Makoto Matsumoto             */
'    mt[mti] &= 0xffffffffUL;
'    /* for >32 bit machines */

Dim tt As Long

mt(0) = (seed And &HFFFFFFFF)
For mti = 1 To Nuplim
    'original expression, rearranged in one line:
    'mt[mti] = (1812433253UL * (mt[mti-1] ^ (mt[mti-1] >> 30)) + mti);

    tt = mt(mti - 1)
    mt(mti) = uAdd(uMult(1812433253, (tt Xor uDiv(tt, kShr30))), mti)
    'innecesary, due to uAdd() and uMult():
    'mt(mti) = mt(mti) And &HFFFFFFFF   '/* for >32 bit machines */
Next

'The following code is not part of the original C code. I apologize for using this not very
'elegant patch, but there is no simpler way to simulate the use of the "static" word in C,
'given that the VBA's "static" word does not behave in a similar way:
mtb = True      'means mt[N] is already initialized
mag01(0) = 0: mag01(1) = MATRIX_A
End Sub     'init_genrand






Public Sub init_by_array(ByRef init_key As Variant, ByVal key_length As Integer)
'void init_by_array(unsigned long init_key[], int key_length)

'/* initialize by an array with array-length */
'/* init_key is the array for initializing keys */
'/* key_length is its length */
'/* slight change for C++, 2004/2/26 */

'int i, j, k;
Dim i As Integer, j As Integer, k As Integer
Dim tt As Long

'init_genrand(19650218UL);
'i=1; j=0;
'k = (N>key_length ? N : key_length);
init_genrand 19650218
i = 1: j = 0
k = IIf((N > key_length), N, key_length)


'for (; k; k--) {
For k = k To 1 Step -1  'while k<>0, that is: while k>0
    'original expression, rearranged in one line:
    'mt[i] = (mt[i] ^ ((mt[i-1] ^ (mt[i-1] >> 30)) * 1664525UL)) + init_key[j] + j;

    tt = mt(i - 1)
    mt(i) = uAdd(uAdd((mt(i) Xor uMult((tt Xor uDiv(tt, kShr30)), 1664525)), init_key(j)), j)

    'mt[i] &= 0xffffffffUL;          /* for WORDSIZE > 32 machines */
    'innecesary, due to uAdd() and uMult():
    'mt(i) = mt(i) And &HFFFFFFFF    '/* for WORDSIZE > 32 machines */

    'i++; j++;
    'if (i>=N) { mt[0] = mt[N-1]; i=1; }
    'if (j>=key_length) j=0;
    i = i + 1: j = j + 1
    If i >= N Then mt(0) = mt(Nuplim): i = 1
    If j >= key_length Then j = 0
Next


'for (k=N-1; k; k--) {
For k = Nuplim To 1 Step -1
    'original expression, rearranged in one line:
    'mt[i] = (mt[i] ^ ((mt[i-1] ^ (mt[i-1] >> 30)) * 1566083941UL)) - i;  /* non linear */

    tt = mt(i - 1)
    mt(i) = uAdd((mt(i) Xor uMult((tt Xor uDiv(tt, kShr30)), 1566083941)), -i)

    'mt[i] &= 0xffffffffUL;          /* for WORDSIZE > 32 machines */
    'innecesary, due to uAdd() and uMult():
    'mt(i) = mt(i) And &HFFFFFFFF    '/* for WORDSIZE > 32 machines */

    'i++;
    'if (i>=N) { mt[0] = mt[N-1]; i=1; }
    i = i + 1
    If i >= N Then mt(0) = mt(Nuplim): i = 1
Next


'mt[0] = 0x80000000UL;   /* MSB is 1; assuring non-zero initial array */
mt(0) = &H80000000      '/* MSB is 1; assuring non-zero initial array */

End Sub     'init_by_array





Public Function genrand_int32SignedLong() As Long   'unsigned long genrand_int32(void)
'This is the translation to VBA of the original C code for genrand_int32(), but renamed
'as explained in the section "Differences with the original C functions and source file"
'/* generates a random number on [0,0xffffffff]-interval */
'(Yes, BUT RETURNS IT AS A (signed) Long in the range [-2^31, 2^31-1])

'unsigned long y;
Dim y As Long

'The below lines were replaced by another approach. See section "On performance" for details:
'static unsigned long mag01[2]={0x0UL, MATRIX_A};
'/* mag01[x] = x * MATRIX_A  for x=0,1 */


If Not mtb Then     'needed in Visual Basic
    'This code is not part of the original C code. It is executed ONLY ONCE in the
    'lifetime of this program. I apologize for using this not very elegant patch,
    'but there is no simpler way to simulate the use of the "static" word in C, given
    'that the VBA's "static" word does not behave in a similar way:

    mti = Nplus1    '/* mti==N+1 means mt[N] is not initialized */
End If


If (mti >= N) Then  '{ /* generate N words at one time */
    'int kk;
    Dim kk As Long

    'if (mti == N+1)   /* if sgenrand() has not been called, */
    '  init_genrand(5489UL); /* a default initial seed is used */
    If mti = Nplus1 Then init_genrand kDefaultSeed

    'for (kk=0;kk<N-M;kk++) {
    '    y = (mt[kk]&UPPER_MASK)|(mt[kk+1]&LOWER_MASK);
    '    mt[kk] = mt[kk+M] ^ (y >> 1) ^ mag01[y & 0x1UL];
    '}
    For kk = 0 To (NuplimLessM)
        y = (mt(kk) And UPPER_MASK) Or (mt(kk + 1) And LOWER_MASK)
        mt(kk) = (mt(kk + M) Xor uDiv(y, kShr1)) Xor mag01(y And &H1)
    Next

    'for (;kk<N-1;kk++) {
    '    y = (mt[kk]&UPPER_MASK)|(mt[kk+1]&LOWER_MASK);
    '    mt[kk] = mt[kk+(M-N)] ^ (y >> 1) ^ mag01[y & 0x1UL];
    '}
    For kk = kk To NuplimLess1
        y = (mt(kk) And UPPER_MASK) Or (mt(kk + 1) And LOWER_MASK)
        mt(kk) = (mt(kk + kDiffMN) Xor uDiv(y, kShr1)) Xor mag01(y And &H1)
    Next

    'y = (mt[N-1]&UPPER_MASK)|(mt[0]&LOWER_MASK);
    'mt[N-1] = mt[M-1] ^ (y >> 1) ^ mag01[y & 0x1UL];
    y = (mt(Nuplim) And UPPER_MASK) Or (mt(0) And LOWER_MASK)
    mt(Nuplim) = (mt(Muplim) Xor uDiv(y, kShr1)) Xor mag01(y And &H1)

    'mti = 0;
    mti = 0
End If


y = mt(mti): mti = mti + 1
'/* Tempering */
'y ^= (y >> 11);
y = (y Xor uDiv(y, kShr11))
'y ^= (y << 7) & 0x9d2c5680UL;
y = (y Xor uMult(y, kShl7) And &H9D2C5680)
'y ^= (y << 15) & 0xefc60000UL;
y = (y Xor uMult(y, kShl15) And &HEFC60000)
'y ^= (y >> 18);
'y = (y Xor uDiv(y, kShr18))    'this step is condensed with the next:
'return y;
genrand_int32SignedLong = (y Xor uDiv(y, kShr18))
End Function    'genrand_int32SignedLong





Public Function genrand_int32() As Double   'unsigned long genrand()
'Returns a value in the range [0, 2^32-1] (that is: [0, 4294967295] )

'WARNINGS:
'   - The return type of the function is Double, not Long, but the values returned are
'     integers.
'   - If you want Long values in the range [-2^31, 2^31-1] ([-2147483648, 2147483647]),
'     then call genrand_int32SignedLong() instead of this function.

Dim tmp As Long

tmp = genrand_int32SignedLong()

If tmp < 0 Then
    genrand_int32 = tmp + k2_32
Else
    genrand_int32 = tmp
End If

End Function    'genrand_int32





Public Function genrand_int31() As Long   'long genrand_int31(void)
'/* generates a random number on [0,0x7fffffff]-interval */
'return (long)(genrand_int32()>>1);
genrand_int31 = CLng(uDiv2(genrand_int32(), kShr1))
End Function    'genrand_int31




Public Function genrand_real1() As Double   'double genrand_real1(void)
'/* generates a random number on [0,1]-real-interval */
'return genrand_int32()*(1.0/4294967295.0);     '/* divided by 2^32-1 */
genrand_real1 = genrand_int32() * kMT_1
End Function    'genrand_real1




Public Function genrand_real2() As Double   'double genrand_real2(void)
'/* generates a random number on [0,1)-real-interval */
'return genrand_int32()*(1.0/4294967296.0);     '/* divided by 2^32 */
genrand_real2 = genrand_int32() * kMT_2
End Function    'genrand_real2




Public Function genrand_real3() As Double   'double genrand_real3(void)
'/* generates a random number on (0,1)-real-interval */
'return (((double)genrand_int32()) + 0.5)*(1.0/4294967296.0);   '/* divided by 2^32 */
genrand_real3 = (genrand_int32() + 0.5) * kMT_3
End Function    'genrand_real3





Public Function genrand_res53() As Double   'double genrand_res53(void)
'/* generates a random number on [0,1) with 53-bit resolution*/
'unsigned long a=genrand_int32()>>5, b=genrand_int32()>>6;
'return(a*67108864.0+b)*(1.0/9007199254740992.0);
genrand_res53 = kMT_res53 * (uDiv2(genrand_int32(), kShr5) * 67108864# + _
                             uDiv2(genrand_int32(), kShr6))
End Function    'genrand_res53


'/* These (PREVIOUS) real versions are due to Isaku Wada, 2002/01/09 added */





'The following functions are present only in the Visual Basic version, not in the
'C version. See more comments in the definition of the constants used as factors:


Public Function genrand_real2b() As Double
'Returns results in the range [0,1) == [0, 1-kMT_Gap2]
'Its lowest value is : 0.0
'Its highest value is: 0.9999999999990
genrand_real2b = genrand_int32() * kMT_2b
End Function    'genrand_real2b


Public Function genrand_real2c() As Double
'Returns results in the range (0,1] == [0+kMT_Gap2, 1.0]
'Its lowest value is : 0.0000000000010  (1E-12)
'Its highest value is: 1.0
genrand_real2c = kMT_Gap2 + (genrand_int32() * kMT_2c)  '==kMT_Gap2+genrand_real2b()
End Function    'genrand_real2c


Public Function genrand_real3b() As Double   'double genrand_real3(void)
'Returns results in the range (0,1) == [0+kMT_Gap, 1-kMT_Gap]
'Its lowest value is : 0.0000000000005  (5E-13)
'Its highest value is: 0.9999999999995
genrand_real3b = kMT_Gap + (genrand_int32() * kMT_3b)
End Function    'genrand_real3b



'Mr. Kenneth C. Ives sent me some code and the idea in which I based genrand_real4b() and
'genrand_real5b(). Added on 2005-Sep-12:


Public Function genrand_real4b() As Double
'Returns results in the range [-1,1] == [-1.0, 1.0]
'Its lowest value is : -1.0
'Its highest value is: 1.0
genrand_real4b = (genrand_int32() * kMT_4b) - 1#
End Function    'genrand_real4b


Public Function genrand_real5b() As Double
'Returns results in the range (-1,1) == [-kMT_GapInterval, kMT_GapInterval]
'Its lowest value is : -0.9999999999990
'Its highest value is: 0.9999999999990
genrand_real5b = kMT_Gap2 + ((genrand_int32() * kMT_5b) - 1#)
End Function    'genrand_real5b







Private Sub main()   'int main(void)
'int i;
'unsigned long init[4]={0x123, 0x234, 0x345, 0x456}, length=4;
Dim i As Integer
Dim init As Variant: init = Array(&H123, &H234, &H345, &H456)
Dim length As Long: length = 4

Const kMaxPrint As Long = 1000 - 1
Dim tmp As Double, s As String


'init_by_array(init, length);
init_by_array init, length


Open "mt19937arVBTest.txt" For Output As #1    'open the output file

'printf("1000 outputs of genrand_int32()\n");
Print #1, Trim(kMaxPrint + 1); " outputs of genrand_int32()"

'for (i=0; i<1000; i++) {
'  printf("%10lu ", genrand_int32());
'  if (i%5==4) printf("\n");
'}
For i = 0 To kMaxPrint
    tmp = genrand_int32()

    s = tmp: If Len(s) < 10 Then s = Space(10 - Len(s)) & s
    Print #1, s;: If i Mod 5 = 4 Then Print #1, " " Else Print #1, " ";
Next


'printf("\n1000 outputs of genrand_real2()\n");
Print #1, "": Print #1, Trim(kMaxPrint + 1); " outputs of genrand_real2()"

'for (i=0; i<1000; i++) {
'  printf("%10.8f ", genrand_real2());
'  if (i%5==4) printf("\n");
'}
For i = 0 To kMaxPrint
    tmp = genrand_real2()

    s = Format(tmp, "0.00000000")
    'to force decimal point instead of decimal comma (as we use in Argentina):
    If tmp < 1# And tmp > 0# Then s = "0." & Right(s, 8)
    Print #1, s;: If i Mod 5 = 4 Then Print #1, " " Else Print #1, " ";
Next
'Print #1, ""


Close #1    'close the output file

'return 0;
End Sub     'main


