########################################################
### Packages/Environment Setup and Functions Created ###

# Data source: https://mavenanalytics.io/challenges/maven-power-outage-challenge/28

# The readxl library will be required for reading the Excel workbook provided for this project.
library(readxl)
# The tidyr library will be required for expanding columns when more than one NERC region is included in a single record.
library(tidyr)
# The odbc and DBI libraries will be required for reading the cleaned data into a SQL Server DB.
library(odbc)
library(DBI)

# Environment variable set for cleaning Excel's numeric time format.
# Excel dates are presented as number strings counting up from December 30th, 1899.
Excel.Date <- '1899-12-30'

## RowAsHeaders(df) is a function to address an issue in the Power Outages dataset where the first row is not the true table header. ##
RowAsHeaders <- function(df){
  if (is.na(df[1,1]) == TRUE | is.na(df[1,2]) == TRUE){
    # Checks if either the first or second value in the first row is NA. 
    # Assumes the header information is in the second row if either condition is TRUE.
    colnames(df) <- df[2,]
    df <- df[-2,]
  } else {
    # Otherwise the function assumes the header information is in the first row.
    colnames(df) <- df[1,]
    df <- df[-1,]  
  }
  return(df)
}

## RemoveRowsNA(df) takes a data frame as input and removes rows that meet a threshold of missing values. ##
# This function was developed to address a problem in the data where several rows contained a single cell of metadata and is designed to remove rows where most column entries are blank.
RemoveRowsNA <- function(df){ 
  df <- subset(df, rowSums(is.na(df)) < ncol(df) - 2)
  return(df)
}

## DatesFromExcel(datecol) converts values from Excel's numeric date format to R Date format. ##
# The process-flow I will utilize when cleaning the dates in this dataset will transform all dates to match Excel's numeric format, then transform them to R Date format.
DatesFromExcel <- function(datecol){return( as.Date(Excel.Date) + as.integer(datecol) )}

## DatesFromExcel(datecol) converts values from %m/%d/%y date format to Excel's Date format. ##
# This function was made to address exceptions in the data that don't conform to Excel's numeric date format.
DatesToExcel <- function(datecol){
  # Retrieves only dates containing a forward slash.
  # The function is designed to ignore dates that are already in Excel's numeric date format.
  filtered_date <- datecol[grepl('/', datecol) == TRUE]
  # Addresses anomalies in the data where multiple forward slashes are used as delimiters.
  # Changes formats from %m/%d/%y to R Date format (%Y-%m-%d).
  formatted_date <- as.Date( gsub('//', '/', filtered_date), format = '%m/%d/%y')
  # Excel dates are presented as number strings counting up from December 30th, 1899.
  # Subtracting a date value from 1899-12-30 yields an integer counting up from the Excel start date.
  to_excel <- as.integer(formatted_date - as.Date(Excel.Date))
  # Assigning this cleaned subset back to the original vector and returning the original vector.
  datecol[grepl('/', datecol) == TRUE] <- to_excel
  return(datecol)
}

## TimesToExcel(timecol) is similar to DatesFromExcel(datecol), but relates to time denoted as a fraction of a day. ##
TimesToExcel <- function(timecol){
  # Removes common variations of times in the dataset.
  # For the purpose of this analysis, we will assume approximate times stated on the DOE form are the start times.
  timecol <- gsub('12:00 noon|12 noon|noon', '12:00 p.m.', timecol, ignore.case = TRUE)
  timecol <- gsub('12:00 midnight|midnight', '12:00 a.m.', timecol, ignore.case = TRUE)
  timecol <- gsub('Approximately ', '', timecol, ignore.case = TRUE)
  # Subsetting only times that appear in the %I:%M %p format.
  filtered_time <- timecol[grepl(':', timecol) == TRUE]
  # Removing periods (.) from times listed as 'p.m.' and 'a.m.' and converting it to a POSIXct time.
  formatted_time <- as.POSIXct( gsub('\\.', '', filtered_time), format = '%I:%M %p', tz = 'UTC')
  # Assigning this cleaned subset back to the original vector and returning the original vector.
  to_excel <- as.numeric(formatted_time - as.POSIXct(Sys.Date()))/(24*60*60)
  timecol[grepl(':', timecol) == TRUE] <- to_excel
  return(timecol)
}

## DateTimeToExcel(datetime, yearcol) converts datetime data to an Excel format, using a year helper-column for entries that do not include a year. ##
DateTimeToExcel <- function(datetime, yearcol) { 
  RestTime <- data.frame(Dates = datetime, Years = yearcol)
  # Removes common variations of times in the dataset.
  # For the purpose of this analysis, we will assume approximate times stated on the DOE form are the start times.
  RestTime$Dates <- gsub('12:00 noon|12 noon|noon', '12:00 pm', RestTime$Dates, ignore.case = TRUE)
  RestTime$Dates <- gsub('12:00 midnight|midnight', '12:00 am', RestTime$Dates, ignore.case = TRUE)
  RestTime$Dates <- gsub('Approximately', '', RestTime$Dates, ignore.case = TRUE)
  # Take a subset of the dates in the %m/%d/%y, %I:%M %p format using the comma as a condition.
  comma.datetimes <- RestTime$Dates[grepl(',', RestTime$Dates) == TRUE]
  # Removing periods (.) from times listed as 'p.m.' and 'a.m.' and converting it to a POSIXct time.  
  comma.datetimes <- gsub('\\.', '', comma.datetimes)
  comma.datetimes <- strptime(comma.datetimes, format='%m/%d/%y, %I:%M %p')
  comma.datetimes <- as.character(as.POSIXct(comma.datetimes, tz = 'UTC') - as.POSIXct(Excel.Date, tz = 'UTC'))
  # Adding this back to the original vector.
  RestTime$Dates[grepl(',', RestTime$Dates) == TRUE] <- comma.datetimes
  # Taking the dates in Time MonthName # format, concatenating the year helper-column, and replacing the month name with its number.
  month.datetimes <- paste0( 
    RestTime$Dates[grepl('ary|ber|March|April|May|June|July|August', RestTime$Dates) == TRUE],
    '/',
    RestTime$Years[grepl('ary|ber|March|April|May|June|July|August', RestTime$Dates) == TRUE] )
  month.list <- c('January ', 'February ', 'March ', 'April ', 'May ', 'June ', 'July ', 'August ', 'September ', 'October ', 'November ', 'December ')
  for (i in 1:length(month.list)) {
    month.datetimes <- gsub(month.list[i]  , paste0(i,'/') , month.datetimes)
  }
  # Removing periods (.) from times listed as 'p.m.' and 'a.m.' and converting it to a POSIXct time.  
  month.datetimes <- gsub('\\.', '', month.datetimes)
  month.datetimes <- strptime(month.datetimes, format='%I:%M %p %m/%d/%Y')
  month.datetimes <- as.character(as.POSIXct(month.datetimes, tz = 'UTC') - as.POSIXct(Excel.Date, tz = 'UTC'))
  # Adding this cleaned subset back to the original vector.
  RestTime$Dates[grepl('ary|ber|March|April|May|June|July|August', RestTime$Dates) == TRUE] <- month.datetimes
  # The remaining datetimes are identified as any column still containing '/' after previous transformations and subsetted from the vector.  
  remaining.datetimes <- RestTime$Dates[grepl('/', RestTime$Dates) == TRUE]
  # Removing periods (.) from times listed as 'p.m.' and 'a.m.' and converting it to a POSIXct time.  
  remaining.datetimes <- gsub('\\.', '',  remaining.datetimes)
  remaining.datetimes <- strptime(remaining.datetimes, format='%m/%d/%y %I:%M %p')
  remaining.datetimes <- as.character(as.POSIXct(remaining.datetimes, tz = 'UTC') - as.POSIXct(Excel.Date, tz = 'UTC'))
  # Assigning this cleaned subset back to the original vector.
  RestTime$Dates[grepl('/', RestTime$Dates) == TRUE] <- remaining.datetimes
  # Returning the formatted original vector.
  RestTime$Dates <- as.numeric(RestTime$Dates)
  
  return(RestTime$Dates)
}

##################################################
### Reading in Data & High Level Data Cleaning ###

# Downloads the zipped data from the link provided for the challenge to a temporary folder.
link <- 'https://maven-datasets.s3.amazonaws.com/Electric+Disturbance+Events/Electric+Disturbance+Events.zip'
temp <- tempfile()
download.file(link, temp)

## Iteratively reads each sheet of the Excel workbook to an R dataframe.
# Saves the names of each sheet in the Excel workbook as a vector.
sheet <- readxl::excel_sheets(unzip(temp, 'DOE_Electric_Disturbance_Events.xlsx'))
# Vector of missing values in the workbook that will be replaced with NAs.
Missing.Values <- c('Unkonwn', 'unkonwn', 'Unknown', 'unknown', 'UNK', 'Ongoing', 'ongoing','NA','N/A', '', ' ')
for (i in sheet) {
  # Unzips the downloaded data and saves each sheet as an R dataframe while replacing missing value strings with NAs.
  # The assign() function is used rather than the assignment operator '<-' to allow for the iterative naming of each dataframe based on the name of the sheet in the Excel workbook.
  assign(paste0('Power.Outages.', i), readxl::read_excel(unzip(temp, 'DOE_Electric_Disturbance_Events.xlsx'), sheet = i, na = Missing.Values))
  # Fixes headers and removes empty rows.
  assign(paste0('Power.Outages.', i), RowAsHeaders(get(paste0('Power.Outages.', i))))
  assign(paste0('Power.Outages.', i), RemoveRowsNA(get(paste0('Power.Outages.', i))))
  # Adds a helper-column for year based on the name of the sheet that will later be used for data cleaning.
  assign(paste0('Power.Outages.', i), cbind(       get(paste0('Power.Outages.', i)), 'Event Year' = i))
}

# Closes the temporary file.
unlink(temp); rm(temp); rm(link)

#####################################################
### Tables From 2002-2010 With The Same Structure ###

# While reviewing the Excel file provided, the various sheets in the workbook have similar structures in three groups: 2002-2010, 2011-2014, and 2015-2023.
# These sheets read into separate dataframes will be cleaned in three parts then appended together once they are arranged in a similar structure.

# First, the column names are fixed among the 2002-2010 dataframes that share a common structure.
colnames(Power.Outages.2003) <- colnames(Power.Outages.2002)
colnames(Power.Outages.2004) <- colnames(Power.Outages.2002)
colnames(Power.Outages.2005) <- colnames(Power.Outages.2002)
colnames(Power.Outages.2006) <- colnames(Power.Outages.2002)
colnames(Power.Outages.2007) <- colnames(Power.Outages.2002)
colnames(Power.Outages.2008) <- colnames(Power.Outages.2002)
colnames(Power.Outages.2009) <- colnames(Power.Outages.2002)
colnames(Power.Outages.2010) <- colnames(Power.Outages.2002)

# The dataframes are then appended together and the original dataframes are removed.
Power.Outages.2002.2010 <- rbind(Power.Outages.2002, Power.Outages.2003, Power.Outages.2004,
                                 Power.Outages.2005, Power.Outages.2006, Power.Outages.2007,
                                 Power.Outages.2008, Power.Outages.2009, Power.Outages.2010)
rm(Power.Outages.2002);rm(Power.Outages.2003);rm(Power.Outages.2004);rm(Power.Outages.2005);rm(Power.Outages.2006);
rm(Power.Outages.2007);rm(Power.Outages.2008);rm(Power.Outages.2009);rm(Power.Outages.2010);

# Transforming dates to uniformly adhere to Excel numeric date format, then to R Date format.
Power.Outages.2002.2010$Date.Event.Began <- DatesFromExcel(DatesToExcel(Power.Outages.2002.2010$Date))
# Removing rows that didn't have a start date listed.
Power.Outages.2002.2010 <- subset(Power.Outages.2002.2010, is.na(Power.Outages.2002.2010$Date.Event.Began) == FALSE )
# Adding the Excel numeric time and date formats to the Excel time start date to arrive at the datetime that the power event began in POSIXct format.
Power.Outages.2002.2010$DateTime.Event.Began <- as.POSIXct(Excel.Date, tz = 'UTC') + (
  (as.numeric(DatesToExcel(Power.Outages.2002.2010$Date)) + 
     as.numeric(TimesToExcel(Power.Outages.2002.2010$Time)) ) * (24*60*60) )
# Restoration datetime column is standardized into Excel numeric datetime format and then added to the Excel time start date to arrive at the datetime that power was restored in POSIXct format.
Power.Outages.2002.2010$DateTime.Restoration <- as.POSIXct(Excel.Date, tz = 'UTC') + DateTimeToExcel(Power.Outages.2002.2010$'Restoration Time', Power.Outages.2002.2010$'Event Year') * (24*60*60)
# The Event Start datetime is subtracted from the Restoration datetime in Excel numeric format to yield the outage length in days.
Power.Outages.2002.2010$Outage.Length.Days <- DateTimeToExcel(Power.Outages.2002.2010$'Restoration Time', Power.Outages.2002.2010$'Event Year') - 
  (as.numeric(DatesToExcel(Power.Outages.2002.2010$Date)) + as.numeric(TimesToExcel(Power.Outages.2002.2010$Time)) )
# Cleaning the Customers Affected column to remove punctuation, text, and deal with values where 'million' was written out.
unique( Power.Outages.2002.2010$'Number of Customers Affected'[is.na(as.numeric(Power.Outages.2002.2010$'Number of Customers Affected')) == TRUE ])
Power.Outages.2002.2010$'Number of Customers Affected'[grepl('mill',Power.Outages.2002.2010$'Number of Customers Affected') == TRUE] <- 
  as.numeric(gsub(' million', '', 
                  Power.Outages.2002.2010$'Number of Customers Affected'[grepl('mill',Power.Outages.2002.2010$'Number of Customers Affected') == TRUE],
                  ignore.case = TRUE)) * 1000000
Power.Outages.2002.2010$Customers.Affected <- as.integer(trimws(gsub('at peak.*|\\([^)]+\\)|Approx\\.|PG&E|\\,|-.*', '', Power.Outages.2002.2010$'Number of Customers Affected', ignore.case = TRUE)))
# Cleaning the Demand Loss column to remove punctuation and text.
unique( Power.Outages.2002.2010$'Loss (megawatts)'[is.na(as.numeric(Power.Outages.2002.2010$'Loss (megawatts)')) == TRUE ])
Power.Outages.2002.2010$Demand.Loss.Mw <- as.integer(trimws(gsub('at peak|peak|Est\\.|Approx\\.|to.*|\\,|-.*', '', Power.Outages.2002.2010$'Loss (megawatts)', ignore.case = TRUE)))
# Creating new columns to rename, these will be cleaned later in the demonstration.
Power.Outages.2002.2010$Disturbance.Type <- Power.Outages.2002.2010$'Type of Disturbance'
Power.Outages.2002.2010$Affected.Area <- Power.Outages.2002.2010$Area
Power.Outages.2002.2010$NERC.Region <- Power.Outages.2002.2010$'NERC Region'
# Preparing an empty description column to match the structure of other tables that will contain values in this column.
Power.Outages.2002.2010$Description <- NA
# Removing original columns.
Power.Outages.2002.2010 <- Power.Outages.2002.2010[,10:ncol(Power.Outages.2002.2010)]

# A similar process is carried out for the dataframes containing years 2011-2014 and 2015-2023.

#####################################################
### Tables From 2011-2014 With The Same Structure ###

# All of the column names were the same in the 2011-2014 bin and didn't need to be changed to match.
# The dataframes from 2011-2014 were appended into one and the original dataframes were removed from the environment.
Power.Outages.2011.2014 <- rbind(Power.Outages.2011, Power.Outages.2012, Power.Outages.2013, Power.Outages.2014)
rm(Power.Outages.2011);rm(Power.Outages.2012);rm(Power.Outages.2013);rm(Power.Outages.2014)

# In this data, the date and the time columns contain a combination of date, time, and datetime values in the Excel numeric format.
# This process mitigates these inconsistencies by iteratively checking if the date column contains a datetime value and then adding the date and time columns if it does not.
for (i in 1:nrow(Power.Outages.2011.2014)) {
  if(Power.Outages.2011.2014$'Date Event Began'[i] != Power.Outages.2011.2014$'Time Event Began'[i] & 
     is.na(Power.Outages.2011.2014$'Date Event Began'[i]) == FALSE){
    Power.Outages.2011.2014$'Date Event Began'[i] <- 
      as.numeric(Power.Outages.2011.2014$'Date Event Began'[i]) + as.numeric(Power.Outages.2011.2014$'Time Event Began'[i])
  }
}
# Removing rows that didn't have a start date listed.
Power.Outages.2011.2014 <- subset(Power.Outages.2011.2014, is.na(Power.Outages.2011.2014$'Date Event Began') == FALSE )
# Transforming dates to uniformly adhere to R Date format.
Power.Outages.2011.2014$Date.Event.Began <- DatesFromExcel(Power.Outages.2011.2014$'Date Event Began')
# Adding the Excel numeric time and date formats to the Excel time start date to arrive at the datetime that the power event began in POSIXct format.
Power.Outages.2011.2014$DateTime.Event.Began <- as.POSIXct(Excel.Date, tz = 'UTC') + as.numeric(Power.Outages.2011.2014$'Date Event Began') * (24*60*60)
# Using the same iterative process for mitigating inconsistent datetime values for the Restoration datetime as was used for the Event Statrt datetime.
for (i in 1:nrow(Power.Outages.2011.2014)) {
  if(Power.Outages.2011.2014$'Date of Restoration'[i] != Power.Outages.2011.2014$'Time of Restoration'[i] & 
     is.na(Power.Outages.2011.2014$'Date of Restoration'[i]) == FALSE){
    Power.Outages.2011.2014$'Date of Restoration'[i] <- 
      as.numeric(Power.Outages.2011.2014$'Date of Restoration'[i]) + as.numeric(Power.Outages.2011.2014$'Time of Restoration'[i])
  }
}
# Restoration datetime column is standardized into Excel numeric datetime format and then added to the Excel time start date to arrive at the datetime that power was restored in POSIXct format.
Power.Outages.2011.2014$DateTime.Restoration <- as.POSIXct(Excel.Date, tz = 'UTC') + as.numeric(Power.Outages.2011.2014$'Date of Restoration') * (24*60*60)
# The Event Start datetime is subtracted from the Restoration datetime in Excel numeric format to yield the outage length in days.
Power.Outages.2011.2014$Outage.Length.Days <- as.numeric(Power.Outages.2011.2014$'Date of Restoration') - as.numeric(Power.Outages.2011.2014$'Date Event Began')
# Cleaning the Customers Affected column to replace 'none' values with 0's and cast to the integer data type.
unique(Power.Outages.2011.2014$'Number of Customers Affected')
Power.Outages.2011.2014$Customers.Affected <- as.integer(gsub('None',0, Power.Outages.2011.2014$'Number of Customers Affected', ignore.case = TRUE))
# Cleaning the Demand Loss column to replace 'none' values with 0's and cast to the integer data type.
unique(Power.Outages.2011.2014$'Demand Loss (MW)')
Power.Outages.2011.2014$Demand.Loss.Mw <- as.integer(gsub('None',0, Power.Outages.2011.2014$'Demand Loss (MW)', ignore.case = TRUE))
# Creating new columns to rename, these will be cleaned later in the demonstration.
Power.Outages.2011.2014$Disturbance.Type <- Power.Outages.2011.2014$'Event Type'
Power.Outages.2011.2014$Affected.Area <- Power.Outages.2011.2014$'Area Affected'
Power.Outages.2011.2014$NERC.Region <- Power.Outages.2011.2014$'NERC Region'
# Preparing an empty description column to match the structure of other tables that will contain values in this column.
Power.Outages.2011.2014$Description <- NA
# Removing original columns.
Power.Outages.2011.2014 <- Power.Outages.2011.2014[,11:ncol(Power.Outages.2011.2014)]

# A similar process is carried out for the dataframe containing years 2015-2023.

#####################################################
### Tables From 2015-2023 With The Same Structure ###

# A year helper-column was added to all dataframes in the beginning, but the data from 2023 already had one. The duplicate column was removed.
Power.Outages.2023 <- Power.Outages.2023[,-1]
# Only the column names from the 2023 data needed to be renamed in the 2015-2023 bin.
colnames(Power.Outages.2023) <- colnames(Power.Outages.2022)

# The dataframes from 2015-2023 were appended into one and the original dataframes were removed from the environment.
Power.Outages.2015.2023 <- rbind(Power.Outages.2015, Power.Outages.2016, Power.Outages.2017,
                                 Power.Outages.2018, Power.Outages.2019, Power.Outages.2020,
                                 Power.Outages.2021, Power.Outages.2022, Power.Outages.2023)
rm(Power.Outages.2015);rm(Power.Outages.2016);rm(Power.Outages.2017);rm(Power.Outages.2018);rm(Power.Outages.2019)
rm(Power.Outages.2020);rm(Power.Outages.2021);rm(Power.Outages.2022);rm(Power.Outages.2023)

# Transforming dates to uniformly adhere to Excel numeric date format, then to R Date format.
Power.Outages.2015.2023$'Date Event Began'[grepl('/',Power.Outages.2015.2023$'Date Event Began') == TRUE] <- 
  as.Date( Power.Outages.2015.2023$'Date Event Began'[grepl('/',Power.Outages.2015.2023$'Date Event Began') == TRUE], 
           format = '%m/%d/%Y') - as.Date(Excel.Date)
Power.Outages.2015.2023$Date.Event.Began <- DatesFromExcel(Power.Outages.2015.2023$'Date Event Began')
# Adding the Excel numeric time and date formats to the Excel time start date to arrive at the datetime that the power event began in POSIXct format.
Power.Outages.2015.2023$DateTime.Event.Began <- as.POSIXct(Excel.Date, tz = 'UTC') + (
  (as.numeric(Power.Outages.2015.2023$'Date Event Began') +
    as.numeric(Power.Outages.2015.2023$'Time Event Began') ) * (24*60*60) )
# Restoration datetime column is standardized into Excel numeric datetime format and then added to the Excel time start date to arrive at the datetime that power was restored in POSIXct format.
Power.Outages.2015.2023$'Date of Restoration'[grepl('/',Power.Outages.2015.2023$'Date of Restoration') == TRUE] <- 
  (as.Date( Power.Outages.2015.2023$'Date of Restoration'[grepl('/',Power.Outages.2015.2023$'Date of Restoration') == TRUE], 
           format = '%m/%d/%Y') - as.Date(Excel.Date)) 
Power.Outages.2015.2023$'Date of Restoration' <- as.numeric(Power.Outages.2015.2023$'Date of Restoration') + 
  as.numeric(Power.Outages.2015.2023$'Time of Restoration')
Power.Outages.2015.2023$DateTime.Restoration <- as.POSIXct(Excel.Date, tz = 'UTC') + Power.Outages.2015.2023$'Date of Restoration' * (24*60*60)
# The Event Start datetime is subtracted from the Restoration datetime in Excel numeric format to yield the outage length in days.
Power.Outages.2015.2023$Outage.Length.Days <- Power.Outages.2015.2023$'Date of Restoration' - 
  (as.numeric(Power.Outages.2015.2023$'Date Event Began') + as.numeric(Power.Outages.2015.2023$'Time Event Began') )
# The Customers Affected column doesn't require further cleaning in this case.
unique(Power.Outages.2015.2023$'Number of Customers Affected')
Power.Outages.2015.2023$Customers.Affected <- as.numeric(Power.Outages.2015.2023$'Number of Customers Affected')
# The Demand Loss column doesn't require further cleaning in this case.
unique(Power.Outages.2015.2023$'Demand Loss (MW)')
Power.Outages.2015.2023$Demand.Loss.Mw <- as.numeric(Power.Outages.2015.2023$'Demand Loss (MW)')
# Creating new columns to rename, these will be cleaned later in the demonstration.
Power.Outages.2015.2023$Disturbance.Type <- Power.Outages.2015.2023$'Event Type'
Power.Outages.2015.2023$Affected.Area <- Power.Outages.2015.2023$'Area Affected'
Power.Outages.2015.2023$NERC.Region <- Power.Outages.2015.2023$'NERC Region'
Power.Outages.2015.2023$Description <- Power.Outages.2015.2023$'Alert Criteria'
# Removing original columns.
Power.Outages.2015.2023 <- Power.Outages.2015.2023[,13:ncol(Power.Outages.2015.2023)]

#########################################################
### Binding Data to One Table and Additional Cleaning ###

# The cleaned data from the three bins of years containing similar table structures are now appended into one dataframe.
Power.Outages <- rbind(Power.Outages.2002.2010, Power.Outages.2011.2014, Power.Outages.2015.2023)
rm(Power.Outages.2002.2010); rm(Power.Outages.2011.2014); rm(Power.Outages.2015.2023)

summary(Power.Outages)

## Adding helper-columns for report KPIs ##
Power.Outages$Incident.ID <- 1:nrow(Power.Outages)
Power.Outages$No.Customers.Affected <- ifelse(Power.Outages$Customers.Affected == 0, TRUE, FALSE)
Power.Outages$No.Demand.Loss <- ifelse(Power.Outages$Demand.Loss.Mw == 0, TRUE, FALSE)

## Correcting Issues with Restoration DateTime Column ##
Power.Outages$DateTime.Restoration[Power.Outages$DateTime.Restoration > Sys.time()]
# A datetime greater than today's date in this column would mean that the original time and date columns that were merged contained two distinct datetime values, breaking the DateTimeToExcel function results. 
# Since we cannot be sure which of the values were correct, we will discard these values entirely
Power.Outages$DateTime.Restoration[Power.Outages$DateTime.Restoration > Sys.time()] <- NA
# There are several Outage Length values that seem to imply that the outage lasted for either a negative amount of time or several thousand days. 
Power.Outages$Outage.Length.Days[(Power.Outages$Outage.Length.Days < 0 | Power.Outages$Outage.Length.Days > 1000) & 
                                   is.na(Power.Outages$Outage.Length.Days) == FALSE]
# This would imply that there were more issues in the original restoration time values. We will discard these values since we cannot be sure of their true values.
Power.Outages$Outage.Length.Days[Power.Outages$Outage.Length.Days < 0 | Power.Outages$Outage.Length.Days > 1000] <- NA

## Cleaning the Disturbance.Type Column ##
unique(Power.Outages$Disturbance.Type)
# All non-disaster weather related events will be grouped together as Weather Related.
Power.Outages$Disturbance.Type[grepl('ice|weather|severe|storm|snow|rain|wind|lightning|temperature|heat', Power.Outages$Disturbance.Type, ignore.case = TRUE)] <- 'Weather Related'
# All natural disasters will be grouped separate from weather since these events tend to lead to longer outages than regular weather events.
Power.Outages$Disturbance.Type[grepl('tropical|flood|wild|wild fire|tornado|disaster|hurricane|earthquake', Power.Outages$Disturbance.Type, ignore.case = TRUE)] <- 'Natural Disaster'
# All physical damage resulting from crime will be grouped under one category since they often overlap in the original data.
Power.Outages$Disturbance.Type[grepl('vandal|sabotage|physical|theft', Power.Outages$Disturbance.Type, ignore.case = TRUE)] <- 'Physical Attack: Theft, Vandalism, or Sabotage'
# Cyber Event category has been standardized.
Power.Outages$Disturbance.Type[grepl('cyber|telecom', Power.Outages$Disturbance.Type, ignore.case = TRUE)] <- 'Cyber Event'
# All unspecified load shedding events have been grouped together.
Power.Outages$Disturbance.Type[grepl('load|shed', Power.Outages$Disturbance.Type, ignore.case = TRUE)] <- 'Load Shedding'
# All equipment failures resulting from regular operation have been grouped together.
Power.Outages$Disturbance.Type[grepl('generator|lines|fire|failure|trip|fault|island|separation|equipment|malfunction|breaker|switch', Power.Outages$Disturbance.Type, ignore.case = TRUE)] <- 'Equipment Failure'
# All events relating to deficient supply of electricity have been grouped together
Power.Outages$Disturbance.Type[grepl('inadequacy|inadequate|supply|deficiency', Power.Outages$Disturbance.Type, ignore.case = TRUE)] <- 'Supply Deficiency'
# Unknown, Other, and a single incident involving a helicopter colliding with transmission lines have been grouped as Other.
Power.Outages$Disturbance.Type[grepl('other|unk|unknown|heli', Power.Outages$Disturbance.Type, ignore.case = TRUE)] <- 'Other'
# Suspicious Activity category has been standardized.
Power.Outages$Disturbance.Type[grepl('susp', Power.Outages$Disturbance.Type, ignore.case = TRUE)] <- 'Suspicious Activity'
# Regular drills, shutdowns, and interruptions that are consequences of normal operations have been grouped as Regular Service Interruptions.
Power.Outages$Disturbance.Type[grepl('shut down|system|test|drill', Power.Outages$Disturbance.Type, ignore.case = TRUE)] <- 'Regular Service Interruption'
# All other disruptions have been grouped as Unspecified Transmission Disruption. 
Power.Outages$Disturbance.Type[grepl('interruption|reduction|loss|firm|outage|distribution|disruption', Power.Outages$Disturbance.Type, ignore.case = TRUE)] <- 'Unspecified Transmission Disruption'
# All other public appeal events have been grouped as Unspecified Public Appeals.
Power.Outages$Disturbance.Type[grepl('public|appeal', Power.Outages$Disturbance.Type, ignore.case = TRUE)] <- 'Unspecified Public Appeal'

## Cleaning the NERC.Region column ##
unique(Power.Outages$NERC.Region)
# Reliability First (RF) was founded on 1/1/06, and is the successor to three prior reliability organizations: MAAC, ECAR, and MAIN.
Power.Outages$NERC.Region <- gsub('ECAR|MAIN|MAAC|WeEnergiesMAIN|RFC|Midwest ISO \\(RFC|REC', 'RF', Power.Outages$NERC.Region)
# The Southern Power Pool (SPP) and Mid-continent Area Power Pool (MAPP) were dissolved on 5/4/18 to form the Midwest Reliability Organization (MRO).
Power.Outages$NERC.Region <- gsub('SPP|MAPP|MR0', 'MRO', Power.Outages$NERC.Region)
# HI = Hawaii, HECO = Hawaiian Electric, MECO = Maui HECO Affiliate
Power.Outages$NERC.Region <- gsub('HI|MECO', 'HECO', Power.Outages$NERC.Region)
# Consistently formatting delimiters in the NERC Region column.
Power.Outages$NERC.Region <- gsub('\\, |\\; | \\/ |[[:space:]]', '/', Power.Outages$NERC.Region)
# TE, RE, and TRE are all referring to the Texas Reliability Entity (Texas RE). ERCOT is the state-controlled Reliability Council that managers the TRE.
Power.Outages$NERC.Region <- gsub('TRE|RE|TE|ERCOT', 'TRE', Power.Outages$NERC.Region)
# Northeast Power Coordinating Council (NPCC)
Power.Outages$NERC.Region <- gsub('NPCC|NPPC|NP', 'NPCC', Power.Outages$NERC.Region)
# Western Electricity Coordinating Council (WECC)
Power.Outages$NERC.Region <- gsub('WSCC', 'WECC', Power.Outages$NERC.Region)
# The Florida Reliability Coordinating Council (FRCC) is contiguous with the Southeastern Reliability Corporation (SERC).
Power.Outages$NERC.Region <- gsub('FRCC\\/SERC|FRCC', 'SERC', Power.Outages$NERC.Region)
# Correcting NAs in NERC Region column
Power.Outages$NERC.Region[is.na(Power.Outages$NERC.Region) == TRUE & grepl('Puerto Rico', Power.Outages$Affected.Area) == TRUE] <- 'PR' 
Power.Outages$NERC.Region[is.na(Power.Outages$NERC.Region) == TRUE & grepl('Hawaii', Power.Outages$Affected.Area) == TRUE] <- 'HECO'
Power.Outages$NERC.Region[is.na(Power.Outages$NERC.Region) == TRUE & grepl('Joplin', Power.Outages$Affected.Area) == TRUE] <- 'MRO'
# Uses a function found in the tidyr package to split rows when they contain more than 1 NERC region.
# New rows created by this will share Incident IDs with the original row to avoid the double-counting of Demand Loss and Customers
Power.Outages <- tidyr::separate_longer_delim(data = Power.Outages, cols = NERC.Region, delim = '/')

###########################################
### Writing the Data to a SQL Server DB ###

# Database Storage: R will extract the data from the web source and store it in an MS SQL Server DB.
# Access the SQL Server DB.
db <- dbConnect(odbc(), Driver = 'SQL Server', Server = 'localhost\\SQLEXPRESS', Database = 'test_env')

# Write the R dataframe to a SQL Server table.
dbCreateTable(db, 'Power.Outages', Power.Outages)
dbWriteTable(db, 'Power.Outages', Power.Outages, overwrite = TRUE, append = FALSE)

# Query data that was just written to test.
dbReadTable(db, 'Power.Outages')

# Disconnect from DB.
dbDisconnect(db);rm(db)