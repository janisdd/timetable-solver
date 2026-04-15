/* min: 10000 b + rde + rma; */
/* min: rde + rma; */
min: 10000 b + rde + rma;


/* missing subjects */
mde = 2;
mma = 2;

/* b is the difference of the missing subjects */
b + rde - rma >= 0;
b - rde + rma >= 0;

/* take only max missing subjects (not necessary? because rde??)*/
/*
mde - sd1 - sd2 - sd3 >= 0;
mma - sm1 - sm2 - sm3 >= 0;
*/

/* introduce rde as temporary variable */
/* 6 missing deu = */
/* rde is slack (rest after finish) */
/* sd1, ... subject deu 1 -> out var how many we have teacher, subject, class comb */
mde - sd1 - sd2 - sd3 - rde >= 0;
mde - sd1 - sd2 - sd3 - rde <= 0;

1 - sd1 - sd2 - sd3 >= 0;

/* introduce rma as temporary variable */
mma - sm1 - sm2 - sm3 - rma >= 0;
mma - sm1 - sm2 - sm3 - rma <= 0;

/* just a test
1 - sd1 - sm1 >= 0;
1 - sd1 - sm2 >= 0;
1 - sd1 - sm3 >= 0;
*/

/* we have only 3 slots for deu und 3 for math */
bin sd1,sd2,sd3,sm1,sm2,sm3;

/*
beispiel mit
mde = 6;
mma = 1;

es kommt raus
b                               2
rde                             3
rma                             1
mde                             6
mma                             1
sd1                             1
sd2                             1
sd3                             1
sm1                             0
sm2                             0
sm3                             0

mathe wird gar nicht genommen, erst alles deutsch und diff soll möglichst klein werden, also ma senken wäre schlecht.
Wir haben aber noch kapazitäten für mathe, also wäre schon gut, wenn das doch genommen wird

mde = 10;
mma = 8;

b                               0
rde                             7
rma                             7
mde                            10
mma                             8
sd1                             1
sd2                             1
sd3                             1
sm1                             0
sm2                             0
sm3                             1

3x deu und 1x mathe wird genommen, bei mma=7 würde 0x mathe genommen
für das angleichen ist es gut, aber kapazitäten werden nicht voll ausgenutzt
2. run, dann mit angepassten missings und dann weniger sdX, smY ?

*/

