#Preamble
install.packages("xlsx")
install.packages("gdata")
library(xlsx)
library(gdata)

#create empty data frame for each game between 2004-2021
tennis=data.frame(Date=as.Date(character()),Surface=character(),Winner=character(),Loser=character(),B365W=numeric(),B365L=numeric(),stringsAsFactors=F)
#read xls files into data frame
for (x in c(2004:2012)) {
  df=read.xlsx(paste0("Data/",x,".xls"), sheetIndex=1)
  df=df[,c("Date","Surface","Winner","Loser","B365W","B365L")]
  tennis=rbind(tennis,df)
}  
#read xlxs files into data frame
for (x in 2013:2021) {
  df=read.xlsx(paste0("Data/",x,".xlsx"), sheetIndex=1)
  df=df[,c("Date","Surface","Winner","Loser","B365W","B365L")]
  tennis=rbind(tennis,df)
}
#clean data frame
tennis[,c("Winner","Loser")]=trim(tennis[,c("Winner","Loser")])
tennis=na.omit(tennis)

#note that in the files player names are not consistent thus in the ELO ranking, there are multiple observations for some players, for example Mathieu P.H. & Mathieu P..

#create empty data frame for ELO figures for each player
elo=data.frame(Player=character(),Rating=numeric(),Matches=numeric(),K=numeric(),Pwin=numeric(),stringsAsFactors=F)
#set pre-match figures
for (j in c("Winner","Loser")) {
  for (i in 1:nrow(tennis)) {
    if (any(elo[,"Player"]==tennis[i,j])) {
      new=data.frame(Player=character(),Rating=numeric(),Matches=numeric(),K=numeric(),Pwin=numeric(),stringsAsFactors=F)
      elo=rbind(elo,new)
    } else {
      new=data.frame(Player=tennis[i,j],Rating=1500,Matches=0,K=250/(5^0.4),Pwin=0.5,stringsAsFactors=F)
      elo=rbind(elo,new)
    }
  }
}  
#update ELO figures for each player for each game
for (i in 1:nrow(tennis)) {
  
  #calculate the win probabilities for each game
  tennis$Wpwin[i]=1/(1+10^(( elo[which(elo[,"Player"]==tennis[i,"Loser"]),"Rating"]- elo[which(elo[,"Player"]==tennis[i,"Winner"]),"Rating"])/400))
  tennis$Lpwin[i]=1/(1+10^(( elo[which(elo[,"Player"]==tennis[i,"Winner"]),"Rating"]- elo[which(elo[,"Player"]==tennis[i,"Loser"]),"Rating"])/400))
  
  #update Rating, number of matches and K factor for winning player in ELO data frame
  elo[which(elo[,"Player"]==tennis[i,"Winner"]),"Rating"]=elo[which(elo[,"Player"]==tennis[i,"Winner"]),"Rating"]+elo[which(elo[,"Player"]==tennis[i,"Winner"]),"K"]*(1-elo[which(elo[,"Player"]==tennis[i,"Winner"]),"Pwin"])
  elo[which(elo[,"Player"]==tennis[i,"Winner"]),"Matches"]=elo[which(elo[,"Player"]==tennis[i,"Winner"]),"Matches"]+1
  elo[which(elo[,"Player"]==tennis[i,"Winner"]),"K"]=250/((elo[which(elo[,"Player"]==tennis[i,"Winner"]),"Matches"]+5)^0.4)
  
  #update Rating, number of matches and K factor for losing player in ELO data frame
  elo[which(elo[,"Player"]==tennis[i,"Loser"]),"Rating"]=elo[which(elo[,"Player"]==tennis[i,"Loser"]),"Rating"]+elo[which(elo[,"Player"]==tennis[i,"Loser"]),"K"]*(0-elo[which(elo[,"Player"]==tennis[i,"Loser"]),"Pwin"])
  elo[which(elo[,"Player"]==tennis[i,"Loser"]),"Matches"]=elo[which(elo[,"Player"]==tennis[i,"Loser"]),"Matches"]+1
  elo[which(elo[,"Player"]==tennis[i,"Loser"]),"K"]=250/((elo[which(elo[,"Player"]==tennis[i,"Loser"]),"Matches"]+5)^0.4)
  
  #update player win probabilities in ELO data frame
  elo[which(elo[,"Player"]==tennis[i,"Winner"]),"Pwin"]=1/(1+10^(( elo[which(elo[,"Player"]==tennis[i,"Loser"]),"Rating"]- elo[which(elo[,"Player"]==tennis[i,"Winner"]),"Rating"])/400))
  elo[which(elo[,"Player"]==tennis[i,"Loser"]),"Pwin"]=1/(1+10^(( elo[which(elo[,"Player"]==tennis[i,"Winner"]),"Rating"]- elo[which(elo[,"Player"]==tennis[i,"Loser"]),"Rating"])/400))
} 
#order ELO data frame based on descending ratings
elo=elo[order(-elo[,"Rating"]), ]
rownames(elo)=1:nrow(elo)