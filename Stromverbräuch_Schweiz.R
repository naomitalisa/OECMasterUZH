rm(list=ls())

library(xlsx)
library(tidyr)
library(readxl)
library(zoo)
library(tidyverse)
library(dplyr)
library(qpcR)
library(rgl)
library(scales)
library(ggplot2)

#####################################################################################################################################################################################################################################################################
#####################################################################################################################################################################################################################################################################

YearEnd <- 2024
YearStart <- 2009
wd_path <- "~/Desktop/Energieuebersicht/"
daily_data <- data.frame()                          # Preallocate the space

# Loop over excel files
for (ii in YearStart:YearEnd){
  
  print(paste0("Year = ", ii))
  raw_data <- suppressMessages(readxl:: read_excel(paste0(wd_path,"data/EnergieUebersichtCH-",ii,".xlsx"), sheet = "Zeitreihen0h15"))
  
  raw_df <- raw_data[-1,1:4]                      # Relevant columns and rows

  # Convert to numeric values
  df <- data.frame(raw_df[,1],apply(raw_df[,2:4],2, function(x) as.numeric(as.character(x))))
  
  # Handle time changes: e.g., 02.01.XXXX 00:00 still belongs to the previous day, i.e., 01.01.XXXX
  jj <- 1
  while (jj <= nrow(df)){
    if (substr(df[jj,1],12,16) == "00:00"){
      df[jj,1] <- df[jj-1,1]                        # Assign previous day
    }
    jj <- jj+1
  }
  
  # Get dates
  dates <- unique(substr(df[1:(nrow(df)),1],1,10))          # Extract unique days
  
  # Exclude 29. Feb if leap year
  if ("29.02." %in% substr(dates,1,6)){
    df <- df[-which(substr(df[,1],1,6) == "29.02."),]
    dates <- dates[-which(substr(dates,1,6) == "29.02.")]
  }
  
  # Daily df spanning over all days of the specific year
  daily_tmp <- as.data.frame(matrix(NA, nrow = length(dates), ncol = ncol(df)))
  daily_tmp[,1] <- data.frame(dates)
  colnames(daily_tmp) <- c("Date", "Endverbrauch", "Produktion", "Total Verbrauch")
  
  # Loop over dates
  xx <- 1
  while (xx <= length(dates)){
    tmp <- df[which(substr(df[,1],1,10) == dates[xx]),2:4]      # Temporary store all data entries of one day  
    daily_tmp[xx,2:4] <- colSums(tmp[,])/10^6               # Calculate daily sum in GWh
    
    xx <- xx+1
  }
  
  # Append to final df
  daily_data <- rbind(daily_data,daily_tmp)         # Append to final df
}




#####################################################################################################################################################################################################################################################################
# Rolling 7 day average
MA7d <- daily_data %>%
  mutate(Endverbrauch_MA7d      = rollmean(Endverbrauch, k = 7, fill = NA, align = 'right')) %>%
  mutate(Produktion_MA7d        = rollmean(Produktion, k = 7, fill = NA, align = 'right')) %>%
  mutate(Total_Verbrauch_MA7d = rollmean(`Total Verbrauch`, k = 7, fill = NA, align = 'right')) %>%
  ungroup()
MA7d <- MA7d[,c(1,5:7)]
daily_data_Gallery <- cbind(daily_data, MA7d[,2:4])
 
#####################################################################################################################################################################################################################################################################

NumberOfYears <- YearEnd-YearStart+1                     
MA7d$year     <- substr(MA7d$Date,7,10)                 # Year
allDays       <- unique(substr(daily_data$Date,1,6))      # Days (except 29.02.)

zz <- 2
while (zz <= ncol(daily_data)){
  
  MA7d_formatted <- as.data.frame(matrix(NA, nrow = length(allDays), ncol = NumberOfYears)) # Preallocate
  colnames(MA7d_formatted) <- YearStart:YearEnd
  MA7d_formatted$Day <- allDays
  
  for (tt in YearStart:YearEnd){
    tp <- subset(MA7d[,c(1,zz)], MA7d$year == as.character(tt))  # Selects data of specific year
    yy <- 1
    while (yy <= nrow(tp)){
      ww <- which(MA7d_formatted$Day == as.character(substr(tp$Date,1,6))[yy])  # If days match, store row index
      MA7d_formatted[ww,tt-YearStart+1] <- tp[yy,2]
      yy <- yy+1
    }
  }
  
  MA7d_formatted$Date <- substr(paste0(allDays,tt),1,6)
  MA7d_formatted <- MA7d_formatted[,c("Date", YearStart:YearEnd)]
  
  # Final df with current year
  final_df <- as.data.frame(paste0(allDays,YearEnd))
  colnames(final_df) <- "Date"
  final_df$Y2023 <- MA7d_formatted[,which(colnames(MA7d_formatted) == "2023")]
  final_df$CurrentYear <- MA7d_formatted[,which(colnames(MA7d_formatted) == YearEnd)]

  
  # Calculate mean, max and min of 2009-2021
  final_df$Mean_09_21 <- rowMeans(MA7d_formatted[,2:(ncol(MA7d_formatted)-(1+(YearEnd-2023)))], na.rm = T)
  final_df$Max_09_21 <- apply(MA7d_formatted[,2:(ncol(MA7d_formatted)-(2+(YearEnd-2023)))],1, function(x) max(x, na.rm=T))
  final_df$Min_09_21 <- apply(MA7d_formatted[,2:(ncol(MA7d_formatted)-(3+(YearEnd-2023)))],1, function(x) min(x, na.rm=T))
  
  
  Mean_09_21 <- rep(final_df$Mean_09_21, NumberOfYears)
  Max_09_21 <- rep(final_df$Max_09_21, NumberOfYears)
  Min_09_21 <- rep(final_df$Min_09_21, NumberOfYears)
  
  daily_data_Gallery <- qpcR:::cbind.na(daily_data_Gallery, Mean_09_21, Max_09_21, Min_09_21)
  daily_data_Gallery$Date[(nrow(daily_data_Gallery)-364):nrow(daily_data_Gallery)] <- final_df$Date
  
  # Export Stromendendverbrauch

  if (zz == 2){
    cols = c(1, zz, zz+3, 8:10)
    openxlsx:: write.xlsx(daily_data_Gallery[,cols], file= paste0(wd_path,"Stromendverbrauch_CH.xlsx"), sheetName = "dailyData", rowNames = F)
    openxlsx:: write.xlsx(MA7d_formatted, file = paste0(wd_path,"Stromendverbrauch_CH.xlsx"), sheetName = "Rolling7dMA", rowNames = F, append = T)
    openxlsx:: write.xlsx(final_df, file = paste0(wd_path,"Stromendverbrauch_CH.xlsx"), sheetName = "Graph", rowNames = F, append = T)
  }
  
  # Export Stromproduktion
  if (zz == 3){
    cols = c(1, zz, zz+3, 11:13)
    openxlsx:: write.xlsx(daily_data_Gallery[,cols], file= paste0(wd_path,"Stromproduktion_CH.xlsx"), sheetName = "dailyData", rowNames = F)
    openxlsx:: write.xlsx(MA7d_formatted, file = paste0(wd_path,"Stromproduktion_CH.xlsx"), sheetName = "Rolling7dMA", rowNames = F, append = T)
    openxlsx:: write.xlsx(final_df, file = paste0(wd_path,"Stromproduktion_CH.xlsx"), sheetName = "Graph", rowNames = F, append = T)
  }
  
  # Export Total Stromverbrauch
  if (zz == 4){
    cols = c(1, zz, zz+3, 14:16)
    openxlsx:: write.xlsx(daily_data_Gallery[,cols], file= paste0(wd_path,"Stromverbrauch_CH.xlsx"), sheetName = "dailyData", rowNames = F)
    openxlsx:: write.xlsx(MA7d_formatted, file = paste0(wd_path,"Stromverbrauch_CH.xlsx"), sheetName = "Rolling7dMA", rowNames = F, append = T)
    openxlsx:: write.xlsx(final_df, file = paste0(wd_path,"Stromverbrauch_CH.xlsx"), sheetName = "Graph", rowNames = F, append = T)
  }
  
  zz <- zz + 1

}



# Load necessary libraries
library(ggplot2)
library(dplyr)

# Create the plot

#######################
## Stromendverbrauch ##
#######################

df_1 <- suppressMessages(readxl:: read_excel(paste0(wd_path,"/Stromendverbrauch_CH.xlsx"), sheet = "Graph"))
df_1$Date <- as.Date(df_1$Date, format="%d.%m.%Y")
# Define the plot
plot_1 <- ggplot(df_1, aes(x = Date)) +
  # Shaded area between Min_09_21 and Max_09_21, now labeled "Max/Min 2009 bis 2021"
  geom_ribbon(aes(ymin = Min_09_21, ymax = Max_09_21, fill = "Max/Min 2009 bis 2021"), alpha = 0.5) +
  # Line for Y2023 (now labeled as 2023)
  geom_line(aes(y = Y2023, color = "2023"), size = 1) +
  # Line for CurrentYear (now labeled as 2024)
  geom_line(aes(y = CurrentYear, color = "2024"), size = 1) +
  # Line for Mean_09_21, now labeled "Mittelwert 2009 bis 2021"
  geom_line(aes(y = Mean_09_21, color = "Mittelwert 2009 bis 2021"), size = 1) +
  # Customize fill and color to share the same legend
  scale_fill_manual(values = c("Max/Min 2009 bis 2021" = "grey80"), guide = "legend") +  # Grey shaded area
  scale_color_manual(values = c("2024" = "black",
                                "2023" = "#1f78b4",  
                                "Mittelwert 2009 bis 2021" = "#a6cee3"),
                     breaks = c("2024", "2023", "Mittelwert 2009 bis 2021")) +
  # Use guides to merge the legends into one
  guides(
    fill = guide_legend(override.aes = list(color = NA)),  # Hide color from fill legend
    color = guide_legend(override.aes = list(fill = NA))   # Hide fill from color legend
  ) +
  # Format x-axis to show only the month name
  scale_x_date(labels = date_format("%b"), expand = c(0, 0)) +
  # Add labels and title
  labs(title = "Stromendverbrauch Schweiz",
       subtitle = "Gleitender 7-Tage-Durchschnitt",
       x = "",
       y = "GWh",  # Label for Y-axis
       color = "Legend", 
       fill = "Legend") +  # Legend for the fill (shaded area)
  # Customize theme
  theme_minimal() +
  theme(
    # Move the legend inside the plot
    legend.position = c(0.75, 0.85), # Adjust the position as needed
    legend.background = element_rect(fill = "white", color = "black", size = 0.5), # Add a box around the legend
    legend.title = element_blank(), # Remove legend title
    legend.key.size = unit(0.5, 'cm'),  # Adjust key size to make the legend smaller
    legend.text = element_text(size = 8),  # Reduce text size for the legend
    legend.spacing = unit(0.2, 'cm'),  # Reduce the spacing between legend items
    legend.box.margin = margin(5, 5, 5, 5)  # Add margin inside the legend box
  )
plot_1
# Save the plot as a PNG in the specified folder
ggsave(filename = "plot_stromendverbrauch.png", 
       plot = plot_1, 
       path = wd_path,
       width = 8, 
       height = 6, 
       dpi = 300,
       bg = "white")

###################
## Stromverbrauch##
###################

df_2 <- suppressMessages(readxl:: read_excel(paste0(wd_path,"/Stromverbrauch_CH.xlsx"), sheet = "Graph"))
df_2$Date <- as.Date(df_2$Date, format="%d.%m.%Y")
# Define the plot
plot_2 <- ggplot(df_2, aes(x = Date)) +
  # Shaded area between Min_09_21 and Max_09_21, now labeled "Max/Min 2009 bis 2021"
  geom_ribbon(aes(ymin = Min_09_21, ymax = Max_09_21, fill = "Max/Min 2009 bis 2021"), alpha = 0.5) +
  # Line for Y2023 (now labeled as 2023)
  geom_line(aes(y = Y2023, color = "2023"), size = 1) +
  # Line for CurrentYear (now labeled as 2024)
  geom_line(aes(y = CurrentYear, color = "2024"), size = 1) +
  # Line for Mean_09_21, now labeled "Mittelwert 2009 bis 2021"
  geom_line(aes(y = Mean_09_21, color = "Mittelwert 2009 bis 2021"), size = 1) +
  # Customize fill and color to share the same legend
  scale_fill_manual(values = c("Max/Min 2009 bis 2021" = "grey80"), guide = "legend") +  # Grey shaded area
  scale_color_manual(values = c("2024" = "black",
                                "2023" = "#1f78b4",  
                                "Mittelwert 2009 bis 2021" = "#a6cee3"),
                     breaks = c("2024", "2023", "Mittelwert 2009 bis 2021")) +
  # Use guides to merge the legends into one
  guides(
    fill = guide_legend(override.aes = list(color = NA)),  # Hide color from fill legend
    color = guide_legend(override.aes = list(fill = NA))   # Hide fill from color legend
  ) +
  # Format x-axis to show only the month name
  scale_x_date(labels = date_format("%b"), expand = c(0, 0)) +
  # Add labels and title
  labs(title = "Stromverbrauch Schweiz",
       subtitle = "Gleitender 7-Tage-Durchschnitt",
       x = "",
       y = "GWh",  # Label for Y-axis
       color = "Legend", 
       fill = "Legend") +  # Legend for the fill (shaded area)
  # Customize theme
  theme_minimal() +
  theme(
    # Move the legend inside the plot
    legend.position = c(0.75, 0.85), # Adjust the position as needed
    legend.background = element_rect(fill = "white", color = "black", size = 0.5), # Add a box around the legend
    legend.title = element_blank(), # Remove legend title
    legend.key.size = unit(0.5, 'cm'),  # Adjust key size to make the legend smaller
    legend.text = element_text(size = 8),  # Reduce text size for the legend
    legend.spacing = unit(0.2, 'cm'),  # Reduce the spacing between legend items
    legend.box.margin = margin(5, 5, 5, 5)  # Add margin inside the legend box
  )
plot_2

# Save the plot as a PNG in the specified folder
ggsave(filename = "plot_stromverbrauch.png", 
       plot = plot_2, 
       path = wd_path,  # Replace with your desired folder path
       width = 8, 
       height = 6, 
       dpi = 300,
       bg = "white")

#####################
## Stromproduktion ##
#####################

df_3 <- suppressMessages(readxl:: read_excel(paste0(wd_path,"/Stromproduktion_CH.xlsx"), sheet = "Graph"))
df_3$Date <- as.Date(df_3$Date, format="%d.%m.%Y")
# Define the plot
plot_3 <- ggplot(df_3, aes(x = Date)) +
  # Shaded area between Min_09_21 and Max_09_21, now labeled "Max/Min 2009 bis 2021"
  geom_ribbon(aes(ymin = Min_09_21, ymax = Max_09_21, fill = "Max/Min 2009 bis 2021"), alpha = 0.5) +
  # Line for CurrentYear (now labeled as 2024)
  geom_line(aes(y = CurrentYear, color = "2024"), size = 1) +
  # Line for Y2023 (now labeled as 2023)
  geom_line(aes(y = Y2023, color = "2023"), size = 1) +
  # Line for Mean_09_21, now labeled "Mittelwert 2009 bis 2021"
  geom_line(aes(y = Mean_09_21, color = "Mittelwert 2009 bis 2021"), size = 1) +
  # Customize fill and color to share the same legend
  scale_fill_manual(values = c("Max/Min 2009 bis 2021" = "grey80"), guide = "legend") +  # Grey shaded area
  scale_color_manual(values = c("2024" = "black",
                                "2023" = "#1f78b4",  
                                "Mittelwert 2009 bis 2021" = "#a6cee3"),
                     breaks = c("2024", "2023", "Mittelwert 2009 bis 2021")) +
  
  # Use guides to merge the legends into one
  guides(
    fill = guide_legend(override.aes = list(color = NA)),  # Hide color from fill legend
    color = guide_legend(override.aes = list(fill = NA))   # Hide fill from color legend
  ) +
  # Format x-axis to show only the month name
  scale_x_date(labels = date_format("%b"), expand = c(0, 0)) +
  # Add labels and title
  labs(title = "Stromproduktion Schweiz",
       subtitle = "Gleitender 7-Tage-Durchschnitt",
       x = "",
       y = "GWh",  # Label for Y-axis
       color = "Legend", 
       fill = "Legend") +  # Legend for the fill (shaded area)
  # Customize theme
  theme_minimal() +
  theme(
    # Move the legend inside the plot
    legend.position = c(0.8, 0.85), # Adjust the position as needed
    legend.background = element_rect(fill = "white", color = "black", size = 0.5), # Add a box around the legend
    legend.title = element_blank(), # Remove legend title
    legend.key.size = unit(0.5, 'cm'),  # Adjust key size to make the legend smaller
    legend.text = element_text(size = 8),  # Reduce text size for the legend
    legend.spacing = unit(0.2, 'cm'),  # Reduce the spacing between legend items
    legend.box.margin = margin(5, 5, 5, 5)  # Add margin inside the legend box
  )
plot_3

# Save the plot as a PNG in the specified folder
ggsave(filename = "plot_stromproduktion.png", 
       plot = plot_3, 
       path = wd_path,  # Replace with your desired folder path
       width = 8, 
       height = 6, 
       dpi = 300,
       bg = "white")
