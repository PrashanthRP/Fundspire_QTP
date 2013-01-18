Option Explicit @@ hightlight id_;_Browser("Fundspire Login").Page("Fundspire Login").WebEdit("j username")_;_script infofile_;_ZIP::ssf1.xml_;_


Dim FolderName, FundName, IndexName, TimeStamp
TimeStamp = GetDateTimeStamp
FolderName = "1_Bulk Edit"
FundName = "TestFund"+"_"+TimeStamp
IndexName = "TestIndex"+"_"+TimeStamp


Call Login()
'Call AddFolder("TestFolder_001")
'Call AddFund(FolderName,FundName)
'Call AddIndex(FolderName,IndexName)

 @@ hightlight id_;_Browser("Fundspire Login").Page("Fundspire").Link("0 Sample Folder")_;_script infofile_;_ZIP::ssf18.xml_;_

Wait(3)




Dim obj_Desc, child, i, InnerText1
Set  obj_Desc = Description.Create
obj_Desc("micclass").value = "Link"
obj_Desc("html tag").value= "A"
obj_Desc("class").value= "x-tree-node-anchor"
child = Browser("Fundspire Login").Page("Fundspire").ChildObjects(obj_Desc)
msgbox child.count


'Browser("Fundspire Login").Page("Fundspire").Link("0_Funds").Click @@ hightlight id_;_Browser("Fundspire Login").Page("Fundspire").Link("0 Funds")_;_script infofile_;_ZIP::ssf16.xml_;_
'Browser("Fundspire Login").Page("Fundspire").Link("0_Funds").Click @@ hightlight id_;_Browser("Fundspire Login").Page("Fundspire").Link("0 Funds")_;_script infofile_;_ZIP::ssf17.xml_;_

Browser("Fundspire Login").Page("Fundspire").Link("1_fOLDER").Click @@ hightlight id_;_Browser("Fundspire Login").Page("Fundspire").Link("1 fOLDER")_;_script infofile_;_ZIP::ssf20.xml_;_
Browser("Fundspire Login").Page("Fundspire").Link("1_fOLDER").Click @@ hightlight id_;_Browser("Fundspire Login").Page("Fundspire").Link("1 fOLDER")_;_script infofile_;_ZIP::ssf21.xml_;_
Browser("Fundspire Login").Page("Fundspire").Image("s_3").Click @@ hightlight id_;_Browser("Fundspire Login").Page("Fundspire").Image("s 3")_;_script infofile_;_ZIP::ssf22.xml_;_
Browser("Fundspire Login").Page("Fundspire").Image("s_3").Click @@ hightlight id_;_Browser("Fundspire Login").Page("Fundspire").Image("s 3")_;_script infofile_;_ZIP::ssf23.xml_;_
Browser("Fundspire Login").Page("Fundspire").Image("s_3").Click @@ hightlight id_;_Browser("Fundspire Login").Page("Fundspire").Image("s 3")_;_script infofile_;_ZIP::ssf24.xml_;_
Browser("Fundspire Login").Page("Fundspire").Image("s_4").Click @@ hightlight id_;_Browser("Fundspire Login").Page("Fundspire").Image("s 4")_;_script infofile_;_ZIP::ssf25.xml_;_
Browser("Fundspire Login").Page("Fundspire").Image("s_5").Click @@ hightlight id_;_Browser("Fundspire Login").Page("Fundspire").Image("s 5")_;_script infofile_;_ZIP::ssf26.xml_;_
