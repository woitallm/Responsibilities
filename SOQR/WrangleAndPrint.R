library(tidyverse)
library(readxl)
library(readr)
library(pwr)


#Prep the Loc file
locs<-read_xlsx("./SOQR/data/LocationConversions.xlsx")
names(locs)<-c("Org","Committee")

#Prep hr file
hr<-read_xlsx("./SOQR/data/HR.xlsx",skip = 12)
names(hr)<-str_remove_all(names(hr),"[[:punct:] ]")
hr<-hr%>%select(PreferredName,BusinessUnitCentralFunction,IsManager)

#Prep the SO file
so<-read_xlsx("./SOQR/data/E86_SOQR.xlsx",skip=2)
names(so)<-str_remove_all(names(so),"[[:punct:] ]")
so<-so%>%fill(SubmissionID,.direction = "down")%>%
  mutate(SecondAxis=str_replace_all(string = SecondAxis,pattern = " - LP",replacement = " Region"),
         ObservedDate=date(ObservedDate),
         Observer=str_remove(string = DirectoryObserverName,pattern = " - .*"))%>%
  select(-DirectoryObserverName)%>%
  arrange(SecondAxis) %>%
  mutate(Dup=duplicated(SubmissionID)) %>%
  filter(Dup==F) %>%
  select(-Dup)%>%
  arrange(SubmissionID)%>%
  left_join(.,hr,by=c("Observer"="PreferredName"))%>%
  filter(IsManager=="Yes", BusinessUnitCentralFunction=="LP")%>%
  mutate(Review_Month=month(ObservedDate,label = T,abbr = F),
         Year=year(ObservedDate),
         Link=paste("https://encompass.enbridge.com/encompassprod/go.aspx?u=/ims/zsaobs&rid=",SubmissionID, "&tm=0&pm=2",sep=""),
         Environ="Field")%>%
  left_join(x=., y=locs,by = c("SecondAxis"="Org"))%>%
  select(SubmissionID, ObservedDate,Review_Month,Year,Committee,Observer,Link,OwnershipOrganization,SecondAxis,Environ)%>%
  filter(!is.na(Committee))
  write_csv(so,"./SOQR/etl-data/SafetyObs.csv")  

#Get counts and sample sizes
samplesize<-function(N, z=1.96){
  round(((1.96^2)*0.5*(1-0.5)/(0.05^2))/(1+((((1.96^2)*0.5*(1-0.5)/(0.05^2))-1)/N)),0)
}

counts_SO <- so%>%
  group_by(Committee)%>%
  summarise(counts=n())%>%
  mutate(npercent=round(samplesize(sum(.$counts))/sum(.$counts)*100,1),
         sample=round(npercent/100*counts,0))
write_csv(counts_SO,"./SOQR/product/Committe-Counts.csv")

#Copy to the converted table
  upload<-data.frame(matrix(ncol=ncol(so)))
  names(upload)<-names(so)
  for(com in unique(counts_SO$Committee)){
    temp_so<-so%>%filter(Committee==com)%>%as.data.frame()
    n<-counts_SO%>%filter(Committee==com)%>%.$sample
    smpl<-sample(x = nrow(temp_so),size=n)
    upload<-bind_rows(upload,temp_so[smpl,])
  }
  upload<-upload%>%filter(!is.na(Committee))

#write to a CSV
  ts<-as.character(now())
  ts<-str_remove_all(ts,"[[:punct:][:cntrl:][a-z]]")
  path<-"./SOQR/product/Upload"
  fnts<-paste(path,"_",ts,".csv",sep = "")
  fn<-paste(path,".csv",sep = "")
  write_csv(upload,fn)
  write_csv(upload,fnts)

