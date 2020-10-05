/* covid 19
 
data source : https://docs.google.com/spreadsheets/d/1wxdWCVvIm7TxQWs4r0EXx5OZEwYBaAoV/edit#gid=415696584
download manual, replace file excel di DATASETRAW, tiap hari utk update

*/
macro define DATASETRAW "D:\Google Drive listiyo.data\! Project Data\1 datasetraw"
macro define FROMSHEET "D:\Google Drive listiyo.data\! Project Data\3 bgetxlsfromsheet"
macro define XLSCSVTOSHEET "D:\Google Drive listiyo.data\! Project Data\3 csvxls_to_sheetableau"
macro define HREADYDS "D:\Google Drive listiyo.data\! Project Data\3 hreadydatastudio"
macro define READYTOANALYZE "D:\Google Drive listiyo.data\! Project Data\3 readytoanalize"
macro define YTEMP "C:\Ytemp"

/*
1. convert exel ke csv untuk connect source data studio
2. clean up and compress data
*/

*START
local shee "Recap-FormRC19 Result-FormRC19 Rekap Kalimantan Gorontalo KALTIM Kalsel MalukuUtara Jabar Jogja BangkaBelitung Jateng Jatim Jakarta Sulut KEPRI Kalbar Sultra Banten Sumut Sumbar Sulsel Sulbar Kalteng" 
clear
foreach a of local shee {
	import excel "$FROMSHEET\My_CovidID.xlsx", sheet("`a'") first all
	di "`a'"
	cap missings dropvars, force
	cap compress
	cap save "$YTEMP\covid_`a'", replace
	clear
}

*cov1 Recap-FormRC19
use "$YTEMP\covid_Recap-FormRC19.dta", clear
gen update = "update! " + c(current_date) + " " +  c(current_time) in 1
missings dropvars, force
*export delimit "$CSVDS\covid_Recap-FormRC19.csv",delimiter(";") replace
export excel "$CSVDS\covid_Results.xlsx", sheet("covid1", modify) firstrow(variable)

*cov2 Result-FormRC19
use "$YTEMP\covid_Result-FormRC19.dta", clear
replace KABKOTA = "isiKABKOTA " + Provinsi if KABKOTA ==""  
*
split TanggalData, p(" ")
split TanggalData1, p("/")
for var TanggalData11 TanggalData12 TanggalData13 : replace X = "0"+X if length(trim(X))==1
gen Tanggal = "20" + TanggalData13 + "-" + TanggalData12 + "-" + TanggalData11
*
split Datesubmitted, p(" ")
split Datesubmitted1, p("/")
for var Datesubmitted11 Datesubmitted12 Datesubmitted13 : replace X = "0"+X if length(trim(X))==1
gen Datesubmit = "20" + Datesubmitted13 + "-" + Datesubmitted13 + "-" + Datesubmitted11 + " " + Datesubmitted2
*
keep ResponseID Datesubmit NamaDataRanger Provinsi KABKOTA Jum* Tanggal SumberData
ren Tanggal TanggalData
ren Datesubmit Datesubmitted
order ResponseID Datesubmitted
*
*missings dropvars, force
gen update = "update! " + c(current_date) + " " +  c(current_time) in 1
*export delimit "$CSVDS\covid_Result-FormRC19.csv",delimiter(";") replace
export excel "$CSVDS\covid_Results.xlsx", sheet("covid2", modify) firstrow(variable)


*cov3 Rekap
use "$YTEMP\covid_Rekap.dta", clear
*
*Berikut kategori umur menurut Depkes RI (2009):
*1) Masa balita : 0-5 tahun
*2) Masa kanak- kanak : 5-11 tahun
*3) Masa remaja awal : 12-16 tahun
*4) Masa remaja akhir : 17-25 tahun
*5) Masa dewasa awal : 26-35 tahun
*6) Masa dewasa akhur : 36-45 tahun
*7) Masa Lansia Awal : 46-55 tahun
*8) Masa lansia akhir : 56-65 tahun
*9) Masa manula : > 65 tahun
replace Usia = subinstr(Usia," tahun","",.)
replace Usia = "" if Usia=="-"
destring Usia, replace
recode Usia (. = 0 "Cek Usia") (0/5= 1 "Balita") (6/11= 2 "Kanak-kanak") (12/16 = 3 "Remaja awal") (17/25 = 4 "Remaja akhir") (26/35 = 5 "Dewasa awal") (36/45 = 6 "Dewasa akhir")  (46/55 = 7 "Lansia awal")  (56/65 = 8 "Dewasa akhir")  (66/135 = 9 "Manula"), gen(Usia_rc)
*
missings dropvars, force
gen update = "update! " + c(current_date) + " " +  c(current_time) in 1
compress
*export delimit "$CSVDS\covid_Rekap.csv",delimiter(";") replace
export excel "$CSVDS\covid_Results.xlsx", sheet("covid3", modify) firstrow(variable)


