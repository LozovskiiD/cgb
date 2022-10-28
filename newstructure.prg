CLOSE DATABASES
USE rasp EXCL
USE datjob IN 0 EXCL
USE tarfond IN 0 EXCL
USE people IN 0 EXCL
ON ERROR DO ersup
ALTER TABLE rasp ADD COLUMN pambsr N(3)
ALTER TABLE rasp ADD COLUMN psmp N(3)
ALTER TABLE rasp ADD COLUMN pmrek N(3)
ALTER TABLE rasp ADD COLUMN pfap N(3)
ALTER TABLE rasp ADD COLUMN psel N(3)
ALTER TABLE rasp ADD COLUMN ppred N(3)

ALTER TABLE rasp ADD COLUMN kfvr N(3,1)
ALTER TABLE rasp ADD COLUMN pkfvr N(4,2)
ALTER TABLE rasp ADD COLUMN vtime N(2)
ALTER TABLE rasp ADD COLUMN vrmain l
ALTER TABLE rasp ADD COLUMN pkf N(5,2)

ALTER TABLE rasp ADD COLUMN night1 C(70)
ALTER TABLE rasp ADD COLUMN night2 C(70)
ALTER TABLE rasp ADD COLUMN night3 C(70)
ALTER TABLE rasp ADD COLUMN night4 C(70)
ALTER TABLE rasp ADD COLUMN night5 C(70)
ALTER TABLE rasp ADD COLUMN night6 C(70)
ALTER TABLE rasp ADD COLUMN night7 C(70)
ALTER TABLE rasp ADD COLUMN night8 C(70)
ALTER TABLE rasp ADD COLUMN night9 C(70)
ALTER TABLE rasp ADD COLUMN night10 C(70)
ALTER TABLE rasp ADD COLUMN night11 C(70)
ALTER TABLE rasp ADD COLUMN night12 C(70)

ALTER TABLE rasp ADD COLUMN persnight N(3)
ALTER TABLE rasp ADD COLUMN hourp N(7,2)
ALTER TABLE rasp ADD COLUMN hourp2 N(7,2)
ALTER TABLE rasp ADD COLUMN persnight2 N(3)
ALTER TABLE rasp ADD COLUMN opdnight N(3)
ALTER TABLE rasp ADD COLUMN ophday N(6,2)
ALTER TABLE rasp ADD COLUMN ophday2 N(6,2)
ALTER TABLE rasp ADD COLUMN optoth N(5)
ALTER TABLE rasp ADD COLUMN oppost N(3)
ALTER TABLE rasp ADD COLUMN oppost2 N(3)
ALTER TABLE rasp ADD COLUMN opzpnorm N(8,2)
ALTER TABLE rasp ADD COLUMN oppers N(3)
ALTER TABLE rasp ADD COLUMN opzph N(8,2)
ALTER TABLE rasp ADD COLUMN opzph2 N(8,2)
ALTER TABLE rasp ADD COLUMN opsumtot N(10,2)
ALTER TABLE rasp ADD COLUMN opnorm N(6,2)
ALTER TABLE rasp ADD COLUMN opdnight2 N(3)
ALTER TABLE rasp ADD COLUMN ntime N(2)
ALTER TABLE rasp ADD COLUMN srtonight N(8,2)

ALTER TABLE rasp ADD COLUMN fete1 C(80)
ALTER TABLE rasp ADD COLUMN fete2 C(80)
ALTER TABLE rasp ADD COLUMN fete3 C(80)
ALTER TABLE rasp ADD COLUMN fete4 C(80)
ALTER TABLE rasp ADD COLUMN fete5 C(80)
ALTER TABLE rasp ADD COLUMN fete6 C(80)
ALTER TABLE rasp ADD COLUMN fete7 C(80)
ALTER TABLE rasp ADD COLUMN fete8 C(80)
ALTER TABLE rasp ADD COLUMN fete9 C(80)
ALTER TABLE rasp ADD COLUMN fete10 C(80)
ALTER TABLE rasp ADD COLUMN fete11 C(80)
ALTER TABLE rasp ADD COLUMN fete12 C(80)
ALTER TABLE rasp ADD COLUMN srtofete N(8,2)
ALTER TABLE rasp ADD COLUMN hfeteday N(2)
ALTER TABLE rasp ADD COLUMN primfete C(37)


ALTER TABLE rasp ADD COLUMN kpeop N(4)
ALTER TABLE rasp ADD COLUMN dotp N(7)
ALTER TABLE rasp ADD COLUMN dzam N(7)
ALTER TABLE rasp ADD COLUMN m1 N(7,2)
ALTER TABLE rasp ADD COLUMN m2 N(7,2)
ALTER TABLE rasp ADD COLUMN m3 N(7,2)
ALTER TABLE rasp ADD COLUMN m4 N(7,2)
ALTER TABLE rasp ADD COLUMN m5 N(7,2)
ALTER TABLE rasp ADD COLUMN m6 N(7,2)
ALTER TABLE rasp ADD COLUMN m7 N(7,2)
ALTER TABLE rasp ADD COLUMN m8 N(7,2)
ALTER TABLE rasp ADD COLUMN m9 N(7,2)
ALTER TABLE rasp ADD COLUMN m10 N(7,2)
ALTER TABLE rasp ADD COLUMN m11 N(7,2)
ALTER TABLE rasp ADD COLUMN m12 N(7,2)
ALTER TABLE rasp ADD COLUMN itday N(7,2)
ALTER TABLE rasp ADD COLUMN srzp N(7,2)
ALTER TABLE rasp ADD COLUMN zpday N(10,2)
ALTER TABLE rasp ADD COLUMN z1 N(8,2)
ALTER TABLE rasp ADD COLUMN z2 N(8,2)
ALTER TABLE rasp ADD COLUMN z3 N(8,2)
ALTER TABLE rasp ADD COLUMN z4 N(8,2)
ALTER TABLE rasp ADD COLUMN z5 N(8,2)
ALTER TABLE rasp ADD COLUMN z6 N(8,2)
ALTER TABLE rasp ADD COLUMN z7 N(8,2)
ALTER TABLE rasp ADD COLUMN z8 N(8,2)
ALTER TABLE rasp ADD COLUMN z9 N(8,2)
ALTER TABLE rasp ADD COLUMN z10 N(8,2)
ALTER TABLE rasp ADD COLUMN z11 N(8,2)
ALTER TABLE rasp ADD COLUMN z12 N(8,2)
ALTER TABLE rasp ADD COLUMN totzp N(10,2)

ALTER TABLE rasp ADD COLUMN kpeopk N(4)
ALTER TABLE rasp ADD COLUMN dkurs N(7)
ALTER TABLE rasp DROP COLUMN dzamk 
ALTER TABLE rasp DROP COLUMN mk1 
ALTER TABLE rasp DROP COLUMN mk2 
ALTER TABLE rasp DROP COLUMN mk3 
ALTER TABLE rasp DROP COLUMN mk4 
ALTER TABLE rasp DROP COLUMN mk5 
ALTER TABLE rasp DROP COLUMN mk6 
ALTER TABLE rasp DROP COLUMN mk7 
ALTER TABLE rasp DROP COLUMN mk8 
ALTER TABLE rasp DROP COLUMN mk9 
ALTER TABLE rasp DROP COLUMN mk10
ALTER TABLE rasp DROP COLUMN mk11
ALTER TABLE rasp DROP COLUMN mk12
ALTER TABLE rasp DROP COLUMN itdayk
ALTER TABLE rasp ADD COLUMN srzpk N(7,2)
ALTER TABLE rasp ADD COLUMN zpdayk N(8,2)
ALTER TABLE rasp ADD COLUMN zk1 N(8,2)
ALTER TABLE rasp ADD COLUMN zk2 N(8,2)
ALTER TABLE rasp ADD COLUMN zk3 N(8,2)
ALTER TABLE rasp ADD COLUMN zk4 N(8,2)
ALTER TABLE rasp ADD COLUMN zk5 N(8,2)
ALTER TABLE rasp ADD COLUMN zk6 N(8,2)
ALTER TABLE rasp ADD COLUMN zk7 N(8,2)
ALTER TABLE rasp ADD COLUMN zk8 N(8,2)
ALTER TABLE rasp ADD COLUMN zk9 N(8,2)
ALTER TABLE rasp ADD COLUMN zk10 N(8,2)
ALTER TABLE rasp ADD COLUMN zk11 N(8,2)
ALTER TABLE rasp ADD COLUMN zk12 N(8,2)
ALTER TABLE rasp DROP COLUMN totzpk 
ALTER TABLE rasp ADD COLUMN zptotk N(10,2)
ALTER TABLE rasp ADD COLUMN ksezotp N(5,2)

ALTER TABLE rasp ADD COLUMN ksekurs N(5,2)
ALTER TABLE rasp ADD COLUMN lokl L

ALTER TABLE rasp ADD COLUMN dst1 N(6,2)
ALTER TABLE rasp ADD COLUMN dst2 N(6,2)
ALTER TABLE rasp ADD COLUMN dst3 N(6,2)
ALTER TABLE rasp ADD COLUMN dst4 N(6,2)
ALTER TABLE rasp ADD COLUMN dst5 N(6,2)
ALTER TABLE rasp ADD COLUMN dst6 N(6,2)
ALTER TABLE rasp ADD COLUMN dst7 N(6,2)
ALTER TABLE rasp ADD COLUMN dst8 N(6,2)
ALTER TABLE rasp ADD COLUMN dst9 N(6,2)
ALTER TABLE rasp ADD COLUMN dst10 N(6,2)
ALTER TABLE rasp ADD COLUMN dst11 N(6,2)
ALTER TABLE rasp ADD COLUMN dst12 N(6,2)
ALTER TABLE rasp ADD COLUMN dsttot N(6,2)
ALTER TABLE rasp ADD COLUMN lkokl L
ALTER TABLE rasp ADD COLUMN dtn N(3)
ALTER TABLE rasp ADD COLUMN dtf N(3)

ALTER TABLE datjob ADD COLUMN kfvr N(3,1)
ALTER TABLE datjob ADD COLUMN pkfvr N(4,2)
ALTER TABLE datjob ADD COLUMN vtime N(2)
ALTER TABLE datjob ADD COLUMN vr1 N(6,2)
ALTER TABLE datjob ADD COLUMN vr2 N(6,2)
ALTER TABLE datjob ADD COLUMN vr3 N(6,2)
ALTER TABLE datjob ADD COLUMN vr4 N(6,2)
ALTER TABLE datjob ADD COLUMN vr5 N(6,2)
ALTER TABLE datjob ADD COLUMN vr6 N(6,2)
ALTER TABLE datjob ADD COLUMN vr7 N(6,2)
ALTER TABLE datjob ADD COLUMN vr8 N(6,2)
ALTER TABLE datjob ADD COLUMN vr9 N(6,2)
ALTER TABLE datjob ADD COLUMN vr10 N(6,2)
ALTER TABLE datjob ADD COLUMN vr11 N(6,2)
ALTER TABLE datjob ADD COLUMN vr12 N(6,2)
ALTER TABLE datjob ADD COLUMN sumvr N(6,2)
ALTER TABLE datjob ADD COLUMN sumvrtot N(7,2)
ALTER TABLE datjob ADD COLUMN fdprn N(12,2)



ALTER TABLE datjob ADD COLUMN pambsr N(3)
ALTER TABLE datjob ADD COLUMN sambsr N(12,2)
ALTER TABLE datjob ADD COLUMN mambsr N(12,2)


ALTER TABLE datjob ADD COLUMN psmp N(3)
ALTER TABLE datjob ADD COLUMN ssmp N(12,2)
ALTER TABLE datjob ADD COLUMN msmp N(12,2)

ALTER TABLE datjob ADD COLUMN pmrek N(3)
ALTER TABLE datjob ADD COLUMN smrek N(12,2)
ALTER TABLE datjob ADD COLUMN mmrek N(12,2)

ALTER TABLE datjob ADD COLUMN pfap N(3)
ALTER TABLE datjob ADD COLUMN sfap N(12,2)
ALTER TABLE datjob ADD COLUMN mfap N(12,2)

ALTER TABLE datjob ADD COLUMN psel N(3)
ALTER TABLE datjob ADD COLUMN ssel N(12,2)
ALTER TABLE datjob ADD COLUMN msel N(12,2)

ALTER TABLE datjob ADD COLUMN pslwork1 N(3)
ALTER TABLE datjob ADD COLUMN sslwork1 N(12,2)
ALTER TABLE datjob ADD COLUMN mslwork1 N(12,2)

ALTER TABLE datjob ADD COLUMN puch N(3)
ALTER TABLE datjob ADD COLUMN such N(12,2)
ALTER TABLE datjob ADD COLUMN much N(12,2)

ALTER TABLE datjob ADD COLUMN primtxt C(50)

ALTER TABLE datjob ADD COLUMN per_date D
ALTER TABLE datjob ADD COLUMN st_per N(2)
ALTER TABLE datjob ADD COLUMN per_sum N(8,2)

ALTER TABLE datjob ALTER COLUMN tokl N(10,2)
ALTER TABLE datjob ALTER COLUMN patt N(5,2)
ALTER TABLE datjob ALTER COLUMN mtokl N(10,2)

ALTER TABLE datjob ADD COLUMN pkf N(5,2)
ALTER TABLE datjob ADD COLUMN hHour N(6,2)
ALTER TABLE datjob ALTER COLUMN sumvr N(7,4)

ALTER TABLE datjob ADD COLUMN ppred N(3)
ALTER TABLE datjob ADD COLUMN spred N(10,2)
ALTER TABLE datjob ADD COLUMN mpred N(10,2)

**--�������
ALTER TABLE datjob ADD COLUMN d1 N(4)
ALTER TABLE datjob ADD COLUMN d2 N(4)
ALTER TABLE datjob ADD COLUMN d3 N(4)
ALTER TABLE datjob ADD COLUMN d4 N(4)
ALTER TABLE datjob ADD COLUMN d5 N(4)
ALTER TABLE datjob ADD COLUMN d6 N(4)
ALTER TABLE datjob ADD COLUMN d7 N(4)
ALTER TABLE datjob ADD COLUMN d8 N(4)
ALTER TABLE datjob ADD COLUMN d9 N(4)
ALTER TABLE datjob ADD COLUMN d10 N(4)
ALTER TABLE datjob ADD COLUMN d11 N(4)
ALTER TABLE datjob ADD COLUMN d12 N(4)
ALTER TABLE datjob ADD COLUMN dtot N(4)

ALTER TABLE datjob ADD COLUMN dst1 N(6,2)
ALTER TABLE datjob ADD COLUMN dst2 N(6,2)
ALTER TABLE datjob ADD COLUMN dst3 N(6,2)
ALTER TABLE datjob ADD COLUMN dst4 N(6,2)
ALTER TABLE datjob ADD COLUMN dst5 N(6,2)
ALTER TABLE datjob ADD COLUMN dst6 N(6,2)
ALTER TABLE datjob ADD COLUMN dst7 N(6,2)
ALTER TABLE datjob ADD COLUMN dst8 N(6,2)
ALTER TABLE datjob ADD COLUMN dst9 N(6,2)
ALTER TABLE datjob ADD COLUMN dst10 N(6,2)
ALTER TABLE datjob ADD COLUMN dst11 N(6,2)
ALTER TABLE datjob ADD COLUMN dst12 N(6,2)
ALTER TABLE datjob ADD COLUMN dsttot N(6,2)

ALTER TABLE datjob ADD COLUMN zp1 N(8,2)
ALTER TABLE datjob ADD COLUMN zp2 N(8,2)
ALTER TABLE datjob ADD COLUMN zp3 N(8,2)
ALTER TABLE datjob ADD COLUMN zp4 N(8,2)
ALTER TABLE datjob ADD COLUMN zp5 N(8,2)
ALTER TABLE datjob ADD COLUMN zp6 N(8,2)
ALTER TABLE datjob ADD COLUMN zp7 N(8,2)
ALTER TABLE datjob ADD COLUMN zp8 N(8,2)
ALTER TABLE datjob ADD COLUMN zp9 N(8,2)
ALTER TABLE datjob ADD COLUMN zp10 N(8,2)
ALTER TABLE datjob ADD COLUMN zp11 N(8,2)
ALTER TABLE datjob ADD COLUMN zp12 N(8,2)
ALTER TABLE datjob ADD COLUMN zptot N(8,2)

**----�����
ALTER TABLE datjob DROP COLUMN dk1 
ALTER TABLE datjob DROP COLUMN dk2 
ALTER TABLE datjob DROP COLUMN dk3 
ALTER TABLE datjob DROP COLUMN dk4 
ALTER TABLE datjob DROP COLUMN dk5 
ALTER TABLE datjob DROP COLUMN dk6 
ALTER TABLE datjob DROP COLUMN dk7 
ALTER TABLE datjob DROP COLUMN dk8 
ALTER TABLE datjob DROP COLUMN dk9 
ALTER TABLE datjob DROP COLUMN dk10
ALTER TABLE datjob DROP COLUMN dk11
ALTER TABLE datjob DROP COLUMN dk12
ALTER TABLE datjob DROP COLUMN dktot

ALTER TABLE datjob DROP COLUMN dstk1 
ALTER TABLE datjob DROP COLUMN dstk2 
ALTER TABLE datjob DROP COLUMN dstk3 
ALTER TABLE datjob DROP COLUMN dstk4 
ALTER TABLE datjob DROP COLUMN dstk5 
ALTER TABLE datjob DROP COLUMN dstk6 
ALTER TABLE datjob DROP COLUMN dstk7 
ALTER TABLE datjob DROP COLUMN dstk8 
ALTER TABLE datjob DROP COLUMN dstk9 
ALTER TABLE datjob DROP COLUMN dstk10 
ALTER TABLE datjob DROP COLUMN dstk11 
ALTER TABLE datjob DROP COLUMN dstk12 
ALTER TABLE datjob DROP COLUMN dsttotk 

ALTER TABLE datjob ADD COLUMN zpk1 N(8,2)
ALTER TABLE datjob ADD COLUMN zpk2 N(8,2)
ALTER TABLE datjob ADD COLUMN zpk3 N(8,2)
ALTER TABLE datjob ADD COLUMN zpk4 N(8,2)
ALTER TABLE datjob ADD COLUMN zpk5 N(8,2)
ALTER TABLE datjob ADD COLUMN zpk6 N(8,2)
ALTER TABLE datjob ADD COLUMN zpk7 N(8,2)
ALTER TABLE datjob ADD COLUMN zpk8 N(8,2)
ALTER TABLE datjob ADD COLUMN zpk9 N(8,2)
ALTER TABLE datjob ADD COLUMN zpk10 N(8,2)
ALTER TABLE datjob ADD COLUMN zpk11 N(8,2)
ALTER TABLE datjob ADD COLUMN zpk12 N(8,2)
ALTER TABLE datjob ADD COLUMN zptotk N(8,2)

ALTER TABLE datjob ADD COLUMN lokl L
ALTER TABLE datjob ADD COLUMN dotp N(3)
ALTER TABLE datjob ADD COLUMN dzam N(3)
ALTER TABLE datjob ADD COLUMN srzp N(8,2)
ALTER TABLE datjob ADD COLUMN zpday N(6,2)
ALTER TABLE datjob ADD COLUMN nrotp N(5)

ALTER TABLE datjob ADD COLUMN lkokl L
ALTER TABLE datjob ADD COLUMN nrkurs N(5)
ALTER TABLE datjob ADD COLUMN dkurs N(3)
ALTER TABLE datjob ADD COLUMN srzpk N(8,2)
ALTER TABLE datjob ADD COLUMN zpdayk N(6,2)
ALTER TABLE datjob ADD COLUMN pol1 N(5)
ALTER TABLE datjob ADD COLUMN pol2 N(5)

SELECT datjob
INDEX ON STR(kodpeop,4)+STR(kp,3)+STR(kd,3)+STR(tr,1) TAG T6

ALTER TABLE tarfond ADD COLUMN logfprn L
ALTER TABLE tarfond ADD COLUMN lognew L
ALTER TABLE tarfond ADD COLUMN formula1 C(100)

ALTER TABLE people ADD COLUMN staj_tar C(10)
ALTER TABLE people ADD COLUMN staj_today C(10)





ON ERROR
***********************************
PROCEDURE ersup