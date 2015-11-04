<%  
DIM strDBuser, strDBpassword, strConnStr

  '************************************************
  '** Connection string to RuleHearing database **
  '************************************************
' DEVELOPMENT DB - October 2015
  strDBuser = "srvrulehearing"
  strDBpassword = "rh3784"
  strConnStr = "Provider=SQLOLEDB; Data Source=OSL1139.verit.dnv.com; Initial Catalog=RuleHearingTest; " & _
        "User ID=" & strDBuser & ";Password=" & strDBpassword

%>

