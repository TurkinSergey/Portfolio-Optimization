
install.packages("quantmod")
install.packages("lpSolveAPI")
install.packages("lpSolve")
install.packages("xlsx")
install.packages("xlsxjars")
install.packages("dplyr")
install.packages("rJava")



##############
#Load Library#
##############

library(quantmod)
library(lpSolveAPI)
library(lpSolve)
library(xlsx)
library(xlsxjars)
library(dplyr)
library(rJava)

##############################
#how much are you investing? #
##############################

investment <- 100000


############################################
#Put in the tickers you want to be analyzed#
############################################

#index
y1 <- "SPY"

#stocks 

x01 = "XOM"
x02 = "MSFT"
x03 = "JNJ"
x04 = "GE"
x05 = "CVX"
x06 = "WFC"
x07 = "PG"
x08 = "JPM"
x09 = "VZ"
x10 = "PFE"
x11 = "T"
x12 = "IBM"
x13 = "MRK"
x14 = "BAC"
x15 = "DIS"
x16 = "ORCL"
x17 = "PM"
x18 = "INTC"
x19 = "SLB"
x20 = "GOOG"
x21 = "AMZN"
x22 = "CSX"
x23 = "PLAY"
x24 = "NFLX"
x25 = "AAPL"
x26 = "MMM"
x27 = "ADBE"
x28 = "AMGN"
x29 = "BLK"
x30 = "CAT"
x31 = "CSCO"
x32 = "CTSH"
x33 = "DRI"
x34 = "DOW"
x35 = "EBAY"
x36 = "F"
x37 = "GS"
x38 = "HD"
x39 = "INTC"
x40 = "NKE"
x41 = "UA"
x42 = "PPL"
x43 = "QRVO"
x44 = "DGX"
x45 = "RAI"
x46 = "ROST"
x47 = "SRCL"
x48 = "TIF"
x49 = "VAR"
x50 = "YHOO"

z <-c(x01, x02, x03, x04, x05, x06, x07, x08, x09, x10,
      x11, x12, x13, x14, x15, x16, x17, x18, x19, x20,
      x21, x22, x23, x24, x25, x26, x27, x28, x29, x30,
      x31, x32, x33, x34, x35, x36, x37, x38, x39, x40,
      x41, x42, x43, x44, x45, x46, x47, x48, x49, x50,
      y1)


#######################################
# create environment to load data into#
#######################################

e <- new.env()

###################
#download the data#
###################
getSymbols(z, 
           from ="2015-01-01", 
           auto.assign = TRUE,
           to = "2015-12-31",
           reload.Symbols = FALSE, 
           warnings = TRUE,
           src = "google",
           symbol.lookup = TRUE) 

###########################################
#clean the data by:                       #
#                                         #
#Putting close prices into one dataframe  #
#                                         #
#convert the dataframe to a matrix        #
#                                         #
#replace NA values with the mean          #
###########################################

ClosePrices <- do.call(merge, lapply(z, function(x) Cl(get(x))))

head(ClosePrices)

ClosePrices.matrix <- as.matrix(ClosePrices)
ClosePrices.matrix

round(ClosePrices.matrix, digits = 2)

cM <- colMeans(ClosePrices.matrix, na.rm=TRUE)
indx <- which(is.na(ClosePrices.matrix), arr.ind = TRUE)
ClosePrices.matrix[indx] <-cM[indx[,2]]


difference <- (diff(ClosePrices.matrix, Lag=1))
difference <- rbind(0, difference)
dailychange <- (difference/ClosePrices.matrix) 
dailychange

####################################################
#Find Beta through the following steps:            #
#                                                  #
#Create a dataset of the covariance                #
#                                                  #
#Pull out the Covar column for the chosen index    #
#                                                  #
#Calculate Beta (Covar(stock,index) / var(index)   #
####################################################


#create covariance matrix
covar <- data.matrix(cov(ClosePrices.matrix))
covar

#extract the covariances against the index
covar.index <- as.data.frame(covar)$SPY.Close

#calculate beta (Covar(stock, index)/var(index))
Beta <- covar.index/tail(covar.index, n=1)
Beta


####################################################
#Find the expected value of each stock to use as   #
#the expected future price                         #
####################################################

#takes the mean of all return means

dailychange1 <-dailychange[,1:ncol(dailychange) -1]

expmeans <- mean(colMeans(dailychange1, na.rm=TRUE))
expreturn <- investment + (expmeans * investment)

expreturn


#########################################
#determine optimal portfolio allocation #
#########################################

#Build an LP to maximize return
#subject to:
  #limit portfolio beta (sum of weighted betas)
#limit total investment in any one stock to 5%






#reward---means 
#risk----- variance/beta
#returns---- investment*mean
#risk minimization -------investment*variancce<u

#max-------rewards
#constriants-----investment*variance<=u
#           -----sum of investment <=1or1000000
#           -----all investments>=0


Beta.new<-Beta[1:length(Beta)-1]
lprec <- make.lp(0, length(Beta.new))
lp.control(lprec, sense="max")
set.objfn(lprec, colMeans(dailychange1, na.rm=TRUE))
add.constraint(lprec, Beta.new, "<=", 3.25)
add.constraint(lprec, Beta.new, ">=", -2.25)
set.bounds(lprec, lower = rep(0, length(Beta.new)), upper = rep(0.05*investment, length(Beta.new)))
add.constraint(lprec, rep(1, length(Beta.new) ), "<=", investment)


lprec
solve(lprec)
get.objective(lprec)
investmentreal<-get.variables(lprec)

sensitivitymeanfrom<-round(get.sensitivity.obj(lprec)$objfrom, digits = 2)
sensitivitymeanstill<-round(get.sensitivity.obj(lprec)$objtill, digits = 2)



lprec1 <- make.lp(0, length(Beta.new))
lp.control(lprec1, sense="minimize")
set.objfn(lprec1, abs(Beta.new))
set.bounds(lprec1, lower = rep(0, length(Beta.new)), upper = rep(0.05*investment, length(Beta.new)))
add.constraint(lprec1, rep(1, length(Beta.new) ), "<=", investment)
add.constraint(lprec1, rep(1, length(Beta.new) ), ">=", 0.9 * investment)

lprec1
solve(lprec1)
get.objective(lprec1)
betasreal<-get.variables(lprec1)

sensitivitybetasfrom<-round(get.sensitivity.obj(lprec1)$objfrom, digits = 2)
sensitivitybetastill<-round(get.sensitivity.obj(lprec1)$objtill, digits = 2)


# all coeffiecents matri
#A = matrix( c(1, 4, 2, 3), nrow = 2, ncol = 2);
A1<-diag(x = 1, length(Beta.new), length(Beta.new))
A<-as.matrix(rbind(Beta.new,Beta.new, rep(1, length(Beta.new) )))
A<-rbind(A,A1)
#Objective fuction
f = c(colMeans(dailychange1, na.rm=TRUE))
#rhs of constriants
b = c( 3.25,  -3.25, investment,rep(0.05*investment, length(Beta.new)))
#signs of constriant
signs = c('<=','>=',"<=","<=")
res = lpSolve::lp('max',f,A,signs,b)
res$solution
get.solutioncount(lprec1)
############################################################################

#get values for 2016 and determine viability  #
###############################################
x<-z
getSymbols(x, 
           from ="2015-01-01", 
           auto.assign = TRUE,
           to = "2016-01-13",
           reload.Symbols = FALSE, 
           warnings = TRUE,
           src = "google",
           symbol.lookup = TRUE) 


ClosePricesx <- do.call(merge, lapply(x, function(x) Cl(get(x))))

head(ClosePricesx)

ClosePricesx.matrix <- as.matrix(ClosePricesx)
ClosePricesx.matrix

round(ClosePricesx.matrix, digits = 2)

cM <- colMeans(ClosePricesx.matrix, na.rm=TRUE)
indx <- which(is.na(ClosePricesx.matrix), arr.ind = TRUE)
ClosePricesx.matrix[indx] <-cM[indx[,2]]


differencex <- (diff(ClosePricesx.matrix, Lag=1))
differencex <- rbind(0, differencex)
dailychangex <- (differencex/ClosePricesx.matrix) 
dailychangex

#create covariance matrix
covarx <- data.matrix(cov(ClosePricesx.matrix))
covarx

#extract the covariances against the index
covarx.index <- as.data.frame(covarx)$SPY.Close

#calculate beta (Covar(stock, index)/var(index))
Betax <- covarx.index/tail(covarx.index, n=1)
Betax<-Betax[1:length(Betax)-1]

dailychangey <-dailychangex[,1:ncol(dailychangex) -1]

expmeansx <- colMeans(dailychangey, na.rm=TRUE)
expmeansx



result<- cbind(colMeans(dailychange1, na.rm=TRUE),sensitivitymeanfrom,expmeansx,sensitivitymeanstill)
resultbeta<- cbind(Beta.new,sensitivitybetasfrom,Betax,sensitivitybetastill)
tickers<-z[1:length(z)-1]
meansss<-colMeans(dailychange1, na.rm=TRUE)

result1<-as.data.frame(result)%>%mutate(decision=ifelse(expmeansx>sensitivitymeanfrom & expmeansx<sensitivitymeanstill,"KEEP","REVIEW"))
result1<-cbind(tickers,result1) 
result2<-as.data.frame(resultbeta)%>%mutate(decision=ifelse(Betax>sensitivitybetasfrom & expmeansx<sensitivitybetastill,"KEEP","REVIEW"))
result2<-cbind(tickers,result2) 

setwd("C:/Users/BravoOne/Documents/Graduate School/Classes/OPR-T680/Final Project/Test Results")

write.xlsx2(covar,"Assignment2.xlsx",sheetName = "covariance",append=TRUE)
write.xlsx2(cbind(meansss,Beta.new),"Assignment2.xlsx",sheetName = "Means",append=TRUE)
write.xlsx2( cbind(z,betasreal),"Assignment2.xlsx",sheetName = "betasmin investment",append=TRUE)
write.xlsx2(cbind(z,investmentreal),"Assignment2.xlsx",sheetName = "meanmax investment",append=TRUE)
write.xlsx2(result1,"Assignment2.xlsx",sheetName = "a",append=TRUE)
write.xlsx2(result2,"Assignment2.xlsx",sheetName = "b",append=TRUE)
