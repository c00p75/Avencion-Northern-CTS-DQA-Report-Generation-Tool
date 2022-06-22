# CTS DQA & Report Genration Tool
This tool is intended to check for CORRECTNESS, CONSISTENCY, and COMPLETENESS of the Avencion retention Client Tracking System (CTS) entries as well generate a summary report of enrollments and harmonization. The tool coverts the .accdb database to an Excel format then applies conditional formatting to indicate any false validation in a particular row and further applies cell borders to falsely validated cells. 

The conditional formatting is applied as follows: 
1.	Green background fill if the entry meets all validations and is complete, consistent, and correct.
2.	Orange background fill if the entry is Not Reviewed or Harmonized, or is Incomplete. It focuses on what makes an entry complete.
3.	Red background fill if the entry is incorrect and inconsistent.

ORANGE BACKGROUND FILL
The row will be flagged orange IF:
•	“Harmonized residential address/Village/Township” = “Not Harmonized”
•	“Harmonized residential address/Village/Township” =  Blank
•	“Harmonized Mobile #” = “Not Harmonized”
•	“Harmonized Mobile #” =  Blank
•	“Enr Type” = “Site Normal Enr” OR “Community Normal Enr” AND “Internal verified” OR “Verfied Mobile No” = “NO”
•	“Next Appointment” = Blank


RED BACKGROUND FILL
The row will be filled red IF:
-	Entry IS NOT one of the dropdown items provided:
-	"Enr Type" != " Site Care Card Enr" & " Site Normal Enr" & " Communty Care Card Enr" & “Community Normal Enr”.
-	"Facility Name" != “Chisanga UHC” & “Kasama General”& “Kaizya HP” & “Kasakalawe HP” & “Mpulungu HAHC” & “Nsumbu RHC” & “Tulemane UHC”
-	“Department” != “MCH” & “Labour Ward” & “VCT” & “PITC” & “DCT” & “Fast Track” & “Traige” & “Youth Conner” & “OPD” & “VMMC” & “Indexing” & “ART” & “T.B” & “Community” & “Cervical Cancer” & “Mens Clinic” &  “Chest Clinic” & “Pediatric Ward” & “PNS”.
-	“Sex” != “Male” &  “Female”.
-	“Status” != “Active” & “Inactive”
-	“Status Comment” != “Trans Out” & “Trans In” & “Deceased” & “LTFU” & “Deactivated” & “Local”
-	“Verfied Mobile No” != “YES” & “NO” & “NO Mobile #”
-	“Internal verified” != “Yes” & “No”
-	“Langueges” != Languages spoken in the region. *
-	“Client Type” != “New” & “Old” 
-	“Client occupation” != “Fishermen/women” & “Farmers” & “Traders” &  “Others”
-	“Harmonized residential address/Village/Township” != “Same in SC” & “Same in PRs” & “Same in Both” &  “Different or No Address in PRs/Added” & “Different or No Address in Sc/Added” & “Different or No Address in Both/Added” & “Same in SC but Different or No Address in PRs” & “Same in PRs but Different or No Address in SC” & “Not Harmonized”.
-	“Address Impacted” != “Yes” &  “No”
-	“Harmonized Mobile #” != “Same in SC” & “Same in PRs” & “Same in Both” & “Different or No Mobile in PRs/Added” & “Different or No Mobile in Sc/Added” & “Different or No Mobile in Both/Added” &  “Same in SC but Different or No Mobile in PRs” & “Same in PRs but Different or No Mobile in SC” & “Care Card” & “Not Harmonized”
-	“VL Harmonization” != “Not Eligible (TX_NEW)” & “Results Found in SC and Updated in CTS” & “Results Found in Physical Registers” & “No VL Result found in SC or PRs” & “VL Updated after follow up” & “VL Results Pending Collection and Updates”

-	Date is Impossible:
•	“Enr date” > Today          
•	“Enr date” < 01/01/2020 *
•	“ArtStartDate” <  1900-01-01
•	“ArtStartDate” > Today
•	"DOB" < 1900-01-01
•	“DOB” > Today
•	“DOB” = Today
•	"Next Appointment" > 7 months (217 days)
•	“Last Appointment” > 7 Months ago ?? 
•	“Revised Next Appointment” < "Enr Date"
•	“Revised Next Appointment” > 7 months (217 days)
•	“VL Due date” > 2 years (730 days)
•	“VL Done Date” > Today
•	“VL Due date” = “VL Done Date”
•	“VL Done Date”  > “ VL Due date”
•	 “Actual Day Seen at Facility” > Today
•	“Status Interaction Date” > Today 
•	“Status Interaction Date” < “Enr date

-	Appointment Management:
•	"Next Appointment" = “Last Appointment” OR “Revised Next Appointment” OR “1 Ap” OR “2 Ap” OR “3 Ap” OR “4 Ap” OR “5 Ap” OR “6 Ap” OR “7 Ap” OR “8 Ap” OR “9 Ap” OR “10 Ap” 
•	"Next Appointment" < “Last Appointment” OR “1 Ap” OR “2 Ap” OR “3 Ap” OR “4 Ap” OR “5 Ap” OR “6 Ap” OR “7 Ap” OR “8 Ap” OR “9 Ap” OR “10 Ap”
•	“Last Appointment” !=  “1 Ap” AND “2 Ap” = BLANK
•	“Last Appointment” !=  “2 Ap” AND “3 Ap” = BLANK
•	“Last Appointment” !=  “3 Ap” AND “4 Ap” = BLANK
•	“Last Appointment” !=  “4 Ap” AND “5 Ap” = BLANK
•	“Last Appointment” !=  “5 Ap” AND “6 Ap” = BLANK
•	“Last Appointment” !=  “6 Ap” AND “7 Ap” = BLANK
•	“Last Appointment” !=  “7 Ap” AND “8 Ap” = BLANK
•	“Last Appointment” !=  “8 Ap” AND “9 Ap” = BLANK
•	“Last Appointment” !=  “9 Ap” AND “10 Ap” = BLANK
•	 “1 Ap” = BLANK  AND “2 Ap” != BLANK
•	 “2 Ap” = BLANK AND “3 Ap” != BLANK
•	 “3 Ap” = BLANK AND “4 Ap” != BLANK
•	 “4 Ap” = BLANK AND “5 Ap”! = BLANK
•	 “5 Ap” =  = BLANK AND “6 Ap” != BLANK
•	 “6 Ap” = BLANK AND “7 Ap” != BLANK
•	 “7 Ap” = BLANK AND “8 Ap” != BLANK
•	 “8 Ap” = BLANK AND “9 Ap” != BLANK
•	 “9 Ap” = BLANK AND “10 Ap” != BLANK

-	Important entry is blank:
•	“Enr date” = Blank
•	“Enr Type” = Blank
•	“Facility Name”= Blank
•	“Department”= Blank
•	“Client Name” = Blank
•	“ART No” = Blank
•	“Client Village/Township” =  Blank
•	“Client Residential Address/Name of Household” =  Blank
•	“DOB (dd/mm/yy)” =  Blank
•	“Sex” =  Blank
•	“Status” =  Blank
•	“Language” = Blank
•	“Enrolled By” = Blank
•	“Status Comment” =  Blank
•	“Status Interaction Date” =  Blank
•	“Client Type” =  Blank
•	 “Internal verified” =  Blank
•	“Verfied Mobile No” =  Blank
•	“VL Harmonization” =  Blank
•	“Address Impacted” =  Blank

-	Status comment is inconsistent with status:
•	"Status" = "Active" AND “Status Comment” = “Trans Out”
•	"Status" = "Active" AND “Status Comment” = “Deceased”
•	"Status" = "Active" AND “Status Comment” = “LTFU”
•	"Status" = "Active" AND “Status Comment” = “Deactivated”
•	“Status” = “Inactive” AND “Status Comment” = “Local”
•	“Status” = “Inactive” AND “Status Comment” = “Trans In”

-	Entry does not match site assessment submission:
•	"Department “ = “Community” AND "Enr Type" = "Site Care Card Enr" OR “Site Normal Enr”
•	“Department” != “Community” AND "Enr Type"  = "Communty Care Card Enr" OR “Community Normal Enr” 

-	Care card enrollment entry is inconsistent: 
•	"Enr Type" = "Site Care Card Enr" AND “Verfied Mobile No” ! = “NO Mobile #”
•	"Enr Type" = "Communty Care Card Enr" AND “Verfied Mobile No” != “NO Mobile #”	
•	"Enr Type" = "Site Care Card Enr" AND “Internal verified” = “Yes”
•	"Enr Type" = "Communty Care Card Enr" AND “Internal verified” = “Yes”
•	"Enr Type" = "Site Care Card Enr" OR “Communty Care Card Enr” AND “Airtel” != Blank
•	"Enr Type" = "Site Care Card Enr" OR “Communty Care Card Enr” AND “Zamtel” != Blank
•	"Enr Type" = "Site Care Card Enr" OR “Communty Care Card Enr”" AND “MTN” != Blank
•	"Enr Type" = "Site Care Card Enr" OR “Communty Care Card Enr” AND “Harmonized Mobile #” != “Care Card”
•	"Enr Type" = "Site Care Card Enr" OR “Communty Care Card Enr” AND “Harmonized Mobile #” != “Not Harmonized”
-	Normal enrollment entry is inconsistent: 
•	"Enr Type" = "Site Normal Enr" OR "Community Normal Enr"  AND “Verfied Mobile No” = “NO Mobile #”
•	"Enr Type" = "Site Normal Enr" OR “Community Normal Enr” AND “Harmonized Mobile #” = “Care Card”
•	"Enr Type" = "Site Normal Enr " OR "Community Normal Enr" AND “Airtel” = Blank
•	"Enr Type" = "Site Normal Enr" OR "Community Normal Enr" AND “Zamtel” = Blank
•	"Enr Type" = "Site Normal Enr" OR "Community Normal Enr" AND “MTN” = Blank
•	"Enr Type" = "Site Normal Enr" OR "Community Normal Enr" AND “Harmonized Mobile #” =  “Care Card”


-	Client type data is inconsistent:
•	“Client Type” = “Old” AND “Enr date” = “ArtStartDate”

-	VL entry is inconsistent:
•	" VL Harmonization " = " Results Found in SC and Updated in CTS" OR" Results Found in Physical Registers" AND “Initial VL” = Blank OR “Current VL” = Blank
•	" VL Harmonization " = " VL Results Pending Collection and Updates " AND “Date VL done” = BLANK

-	Mobile number is incorrect or inconsistent:
•	“Airtel” does not start with “97” OR “77”
•	“Mtn” does not start with “96” OR “76”
“Zamtel” does not start with “95” or “75”
-	Entries associated with residential address are inconsistent:
•	“Harmonized residential address/Village/Township”  = “Not Harmonized” AND “Address Impacted” = “Yes”
•	Address Impacted” = “Yes”  AND “Harmonized residential address/Village/Township” = “Same in SC” OR “Same in PRs” OR “Same in Both” AND “ OR “Not Harmonized” OR BLANK
•	“Harmonized residential address/Village/Township”! = “Same in SC” AND “Same in PRs” AND “Same in Both” AND “Address Impacted” = “No”


GREEN BACKGROUND FILL
The green background fill is programmed to be the opposite of the red background fill. The row will only be flagged green if it does not violet any of the constraints defined above.

WHITE BACKGROUND FILL
No background fill will be applied in cases when no condition has been violated and not all of the conditions have been fulfilled. This will enable the user to identify when the program is not working as it should or needs to be updated.
