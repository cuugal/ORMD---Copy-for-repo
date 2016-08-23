<%
Dim eq
Set eq=Server.CreateObject("Scripting.Dictionary")
eq.Add "Eye Wash Station",       "/pictograms/equipment/eyewash.png"
eq.Add "Safety Shower",          "/pictograms/equipment/shower.png"
eq.Add "Chemical Spill Kit",     "/pictograms/equipment/spill.png"


Dim ppe
Set ppe=Server.CreateObject("Scripting.Dictionary")
ppe.Add "Safety Glasses/Goggles",       "/pictograms/protection/eye.png"
ppe.Add "Face Shield",                    "/pictograms/protection/face.png"
ppe.Add "Safety Footwear",                  "/pictograms/protection/foot.png"
ppe.Add "Hair",                             "/pictograms/protection/hair.png"
ppe.Add "Gloves",                            "/pictograms/protection/hand.png"
ppe.Add "Hard Hat",                         "/pictograms/protection/head.png"
ppe.Add "Hearing Protection",               "/pictograms/protection/hearing.png"
ppe.Add "Protective Clothing/Apron/Overalls",          "/pictograms/protection/ppe.png"
ppe.Add "Respirator/Dust Mask",           "/pictograms/protection/respiratory.png"

%>