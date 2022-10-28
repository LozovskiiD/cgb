CLOSE TABLES
*SET PATH TO ./real
ON ERROR DO erSup
USE rasp EXCL
ALTER TABLE rasp ADD COLUMN pzdrav N(3)
USE datjob EXCL
ALTER TABLE datjob ADD COLUMN pzdrav N(3)
ALTER TABLE datjob ADD COLUMN szdrav N(12,2)
ALTER TABLE datjob ADD COLUMN mzdrav N(12,2)
USE 
ON ERROR 
USE tarfond ORDER 1 IN 0
SELECT tarfond
LOCATE FOR ALLTRIM(LOWER(fpers))=='datjob.pkat'
IF FOUND()
   SCATTER TO a
   APPEN BLANK
   GO TOP 
   GATHER FROM a
   REPLACE rec WITH 'За работу в сфере здравоохранения N 52 п.5',fpers WITH 'datjob.pzdrav',fname WITH 'datjob.szdrav',sayokl WITH 'LTRIM(STR(datjob.szdrav,8,2))', sayoklm WITH 'LTRIM(STR(datjob.mzdrav,8,2))'
   REPLACE procread WITH "DO procTarRead WITH 'datjob.pzdrav','999'", formula WITH 'datJob.mtokl*datJob.pzdrav/100', countvac WITH 'mtokl*pzdrav/100', plrep WITH 'pzdrav',plrepvac WITH 'pzdrav'
   REPLACE firead WITH 'datjob.pzdrav',sum_f WITH 'szdrav',sum_fm WITH 'mzdrav',persved WITH 'pzdrav',sumved WITH 'mzdrav'
   REPLACE formula1 WITH 'datJob.tokl*datJob.pzdrav/100',nameved WITH 'За работу в сфере здравоохранения N 52 п.5'
   BROWS   
ENDIF 
***************************
PROCEDURE erSup
