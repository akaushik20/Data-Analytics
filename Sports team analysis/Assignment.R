
#install and load 'XLConnect' Package
install.packages("XLConnect")
library("XLConnect")

#load the worbook
#Note: The file must be located at the working directory of the system
#We can fetch the Working Directory using "getwd()" command
sample.data<-loadWorkbook("Nike Tournament of Champions Final.xlsx")

#extract each sheet from the workbook
sheet1.wk<-readWorksheet(sample.data, sheet="CHROMO")
sheet2.wk<-readWorksheet(sample.data, sheet="ALPHA")

#CHROMO and ALPHA sheet need to be modified as the column names are different in these sheets
sheet1.wk.mdfy1<-data.frame(DAY=sheet1.wk$DAY,
                            TIME=as.character(sheet1.wk$TIME),
                            COURT=sheet1.wk$COURT,
                            HOME=sheet1.wk$HOME,
                            AWAY=sheet1.wk$AWAY)
sheet2.wk.mdfy<-data.frame(DAY=sheet2.wk$DAY,
                           TIME=as.character(sheet2.wk$TIME),
                           COURT=sheet2.wk$COURT,
                           HOME=sheet2.wk$HOME,
                           AWAY=sheet2.wk$AWAY)


sheet3.wk<-readWorksheet(sample.data, sheet="17U")
sheet4.wk<-readWorksheet(sample.data, sheet="16U")
sheet5.wk<-readWorksheet(sample.data, sheet="15U")
sheet6.wk<-readWorksheet(sample.data, sheet="14U")

#combined all the data extracted from multiple sheets into one dataframe
total.wk<-rbind(sheet1.wk.mdfy1,sheet2.wk.mdfy, sheet3.wk, sheet4.wk,sheet5.wk, sheet6.wk)

#Fetch the "Home" team name and "Away" team name from the data frame 
total.wk.home<-data.frame(total.wk$HOME)
total.wk.away<-data.frame(total.wk$AWAY)

#Rename the column name as TEAMS
#we need to rename the column name because we need to rbind these two columns
colnames(total.wk.home)<-c("TEAMS")
colnames(total.wk.away)<-c("TEAMS")

#combined the two columns
total.teams<-rbind(total.wk.home, total.wk.away)

#Fetch unique Teams name and assign a Unique ID to each row
total.teams.unique<-unique(rbind(total.wk.home, total.wk.away))
total.teams.unique$ID <- 1:nrow(total.teams.unique)

#sort the data frame having unique team name and ID
total.teams.unique.srt<-total.teams.unique[order(total.teams.unique$TEAMS),]

#Remove the NA values from the table
total.teams.unique.srt<- total.teams.unique.srt[-c(693),]

#Add two columns in the master table
#1. Home Team Unique ID
#2. Away Team Unique ID
total.wk$Home_ID<-NA
total.wk$Away_ID<-NA

#Remove the NA values from the table
total.wk<- total.wk[-c(439),]
total.wk<- total.wk[-c(877),]

#Fetch values from the unique team list 
#Add the unique ID for the corrosponding team in the master table
for (i in 1:nrow(total.teams.unique.srt))
{
  team.name<-total.teams.unique.srt$TEAMS[i]
  team.id<-total.teams.unique.srt$ID[i]
  
  for (j in 1:nrow(total.wk))
  {
    
    if (team.name==as.character(total.wk$HOME[j]))
    {
      total.wk$Home_ID[j] <- team.id
    }
    if (team.name==as.character(total.wk$AWAY[j]))
    {
      total.wk$Away_ID[j] <- team.id
    }
  }
}

install.packages("xlsx")
library(xlsx)

write.xlsx(total.wk, "Nike Tournament of Champions Final Sheet.xlsx")
write.xlsx(total.teams.unique.srt, "Unique List of Teams with ID.xlsx")