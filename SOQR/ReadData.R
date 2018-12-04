library(readr)
library(lubridate)

path_HR<-paste("//Cnpl.enbridge.com/common/Edm/SafetyandEnv/SafetyLib/Safety/5.2 Performance Management Monitoring and Reporting/Monthly Reporting/LP Reporting/" ,
               year(rollback(today())) ,
               "/_Current Reporting Period",sep = "")

path_SO<-paste("//Cnpl.enbridge.com/common/Edm/SafetyandEnv/SafetyLib/Safety/5.2 Performance Management Monitoring and Reporting/Safety Observation Quality Reviews/" ,
               year(rollback(today())) ,
               "/_Active Month",sep = "")

file_HR<-paste(path_HR,"HR.xlsx",sep="/")
file_SO<-paste(path_SO,"E86_SOQR.xlsx",sep = "/")
file_Loc<-paste(path_SO,"LocationConversions.xlsx",sep = "/")

try(file.copy(c(file_SO,file_HR,file_Loc),"./SOQR/data"))

