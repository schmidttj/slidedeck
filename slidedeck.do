discard

*  Load and label monthly economic data

insheet using econmo.csv, comma clear
drop date
gen date = mofd(mdy(month, day, year))
format %tmMonYY date
order date
sort date

lab var icpiu "CPI-U index, 1982-84=100, sa"
lab var iip "Industrial production index, 2017=100, sa"
lab var ru3 "U-3 unemployment rate, pct, sa"

*  Create time series graphs

line ru3 date, xtitle("") title("Unemployment Rate")
graph export ru3.png, replace

line iip date, xtitle("") title("Industrial Production")
graph export iip.png, replace

*  Create inflation variable and graph it

gen pcpiu = ((icpiu/icpiu[_n-12])-1)*100
lab var pcpiu "CPI-U inflation, pct, YoY"
line pcpiu date, xtitle("") title("CPI-U Inflation")
graph export pcpiu.png, replace

save econmo, replace

*  Load and label quarterly economic data

insheet using rgdp.csv, comma clear
drop date
gen date = qofd(mdy(month, day, year))
format %tqYY:q date
order date
sort date

lab var rgdp "Real GDP, 2012 dollars, sa"

*  Create real GDP growth variable and graph it

gen rrgdp = ((rgdp/rgdp[_n-4])-1)*100
lab var rrgdp "Real GDP growth, pct, sa"
line rrgdp date, xtitle("") title("Real GDP Growth")
graph export rrgdp.png, replace

save econqtr, replace

*  Create a new deck object

.d = .deck.new "dfsslides.pptx"

*  Create a four-graph slide and add it to the deck object

.s1 = .slide.new econoverview
.s1.set_title "U.S. Economic Conditions"
.s1.add_exhibits rrgdp.png iip.png ru3.png pcpiu.png
.s1.add_margin_bullets - Real GDP continues to expand, but growth is slowing
.s1.add_margin_bullets - Industrial production has recovered strongly from the pandemic-induced recession
.s1.add_margin_bullets - Unemployment remains near its record low given strong labor market conditions
.s1.add_margin_bullets - Consumer price inflation has reached a ** four-decade ** high

.d.add_slide .s1

*  Create an all-text slide and add it to the deck object

.s2 = .slide.new execsummary
.s2.set_title "Executive Summary"
.s2.add_main_bullets + *16 U.S. economic expansion continues, but the pace of growth has slowed
.s2.add_main_bullets ++ Consensus forecast suggests below-trend growth in second half of 2022
.s2.add_main_bullets + *16 Labor market conditions remain very strong amid high labor demand and low participation rates
.s2.add_main_bullets + *16 Unusually high inflation will prompt tighter monetary policy from the Federal Reserve

.d.add_slide .s2

*  Show the slides in the deck

.d.show_slides

*  Delete a slide from the deck and show that it's gone

.d.del_slide .s2
.d.show_slides

*  Add the slide back, this time to the front of the deck

.d.add_slide .s2 1
.d.show_slides

*  Save the deck with the two slides; slides are "rendered" at save

.d.save "c:/users/schmi/documents/git/slidedeck/econreview.pptx"
