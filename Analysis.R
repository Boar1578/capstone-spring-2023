install.packages("readxl") 
install.packages("caTools")
install.packages("Metrics")
# load packages
library(readxl)
library(dplyr)
library(stringr)
library(ggplot2)
library(tidyr)
library(reshape2)
library(forecast)
library(purrr)
library(caTools)
library(Metrics)



# Read in the data from the Excel file
data <- read_excel("/Users/SamBarr7/OneDrive - UW-Eau Claire/Capstone/Commodity Forecasts Consolidated.xlsx", sheet = "Consolidated")
actuals <- read_excel("/Users/SamBarr7/OneDrive - UW-Eau Claire/Capstone/Commodity Actuals.xlsx", sheet = "Export")


# View the first few rows of the data
head(data)

# remove extra top rows
data <- data[-c(1,1),]

#reconfirm
head(data)

# set first row as column names
colnames(data) <- as.character(data[1,])
data <- data[-1,]
head(data)

# remove extra columns
data <- data[, 1:10]
# rename columns for easier coding
names(data)[8] <- "actual_price"
names(data)[9] <- "fcst_price"

# change data type
data$Year <- as.integer(data$Year)
data$Month <- as.integer(data$Month)

data$actual_price <- as.numeric(data$actual_price)
data$fcst_price <- as.numeric(data$fcst_price)

# change so values automatically interpreted as categorical
data <- data %>% 
  mutate(year_month = str_replace_all(Year.Month, "\\.", "_"))

### INDEX METRICS ###
# Group the data by index
grouped_index <- group_by(data, Provider)
grouped_index

# Calculate the Mean Absolute Error (MAE) by index
mae_index <- summarise(grouped_index, MAE = mean(abs(actual_price - fcst_price)))

# Calculate the Mean Squared Error (MSE) by index
mse_index <- summarise(grouped_index, MSE = mean((actual_price - fcst_price)^2))

# Calculate the Root Mean Squared Error (RMSE) by index
rmse_index <- summarise(grouped_index, RMSE = sqrt(mean((actual_price - fcst_price)^2)))

# Calculate the Mean Absolute Percentage Error (MAPE) by index
mape_index <- summarise(grouped_index, MAPE = mean(abs((actual_price - fcst_price)/actual_price)) * 100)

# Print the resulting data frames
print(mae_index)
print(mse_index)
print(rmse_index)
print(mape_index)

### YEAR METRICS ###
# Group the data by year

grouped_year <- group_by(data, Year)
grouped_year

# Calculate the Mean Absolute Error (MAE) by index
mae_year <- summarise(grouped_year, MAE = mean(abs(actual_price - fcst_price)))

# Calculate the Mean Squared Error (MSE) by index
mse_year <- summarise(grouped_year, MSE = mean((actual_price - fcst_price)^2))

# Calculate the Root Mean Squared Error (RMSE) by index
rmse_year <- summarise(grouped_year, RMSE = sqrt(mean((actual_price - fcst_price)^2)))

# Calculate the Mean Absolute Percentage Error (MAPE) by index
mape_year <- summarise(grouped_year, MAPE = mean(abs((actual_price - fcst_price)/actual_price)) * 100)

# Print the resulting data frames
print(mae_year)
print(mse_year)
print(rmse_year)
print(mape_year)

# create the bar chart for RMSE
ggplot(rmse_year, aes(x = Year, y = RMSE)) +
  geom_bar(stat = "identity") + 
  labs(x = "Year", y = "RMSE")


### COMMODITIY METRICS ###
# Group the data by commodity

grouped_comm <- group_by(data, Commodity)
grouped_comm

# Calculate the Mean Absolute Error (MAE) by index
mae_comm <- summarise(grouped_comm, MAE = mean(abs(actual_price - fcst_price)))

# Calculate the Mean Squared Error (MSE) by index
mse_comm <- summarise(grouped_comm, MSE = mean((actual_price - fcst_price)^2))

# Calculate the Root Mean Squared Error (RMSE) by index
rmse_comm <- summarise(grouped_comm, RMSE = sqrt(mean((actual_price - fcst_price)^2)))

# Calculate the Mean Absolute Percentage Error (MAPE) by index
mape_comm <- summarise(grouped_comm, MAPE = mean(abs((actual_price - fcst_price)/actual_price)) * 100)

# Print the resulting data frames
print(mae_comm)
print(mse_comm)
print(rmse_comm)
print(mape_comm)

# create the bar chart for RMSE
ggplot(rmse_comm, aes(x = reorder(Commodity, -RMSE), y = RMSE)) +
  geom_bar(stat = "identity") + 
  labs(x = "Commodity", y = "RMSE")

### REGION METRICS ###
# Group the data by region

grouped_reg <- group_by(data, Region)
grouped_reg

# Calculate the Mean Absolute Error (MAE) by index
mae_reg <- summarise(grouped_reg, MAE = mean(abs(actual_price - fcst_price)))

# Calculate the Mean Squared Error (MSE) by index
mse_reg <- summarise(grouped_reg, MSE = mean((actual_price - fcst_price)^2))

# Calculate the Root Mean Squared Error (RMSE) by index
rmse_reg <- summarise(grouped_reg, RMSE = sqrt(mean((actual_price - fcst_price)^2)))

# Calculate the Mean Absolute Percentage Error (MAPE) by index
mape_reg <- summarise(grouped_reg, MAPE = mean(abs((actual_price - fcst_price)/actual_price)) * 100)

# Print the resulting data frames
print(mae_reg)
print(mse_reg)
print(rmse_reg)
print(mape_reg)

# create the bar chart for RMSE
ggplot(rmse_reg, aes(x = reorder(Region, -RMSE), y = RMSE)) +
  geom_bar(stat = "identity") + 
  labs(x = "Region", y = "RMSE")

### COMMODITIY REGION METRICS ###
# Group the data by commodity_region
grouped_comm_reg <- group_by(data, Commodity, Region)
grouped_comm_reg

# Calculate the Mean Absolute Error (MAE) by index
mae_comm_reg <- summarise(grouped_comm_reg, MAE = mean(abs(actual_price - fcst_price)))

# Calculate the Mean Squared Error (MSE) by index
mse_comm_reg <- summarise(grouped_comm_reg, MSE = mean((actual_price - fcst_price)^2))

# Calculate the Root Mean Squared Error (RMSE) by index
rmse_comm_reg <- summarise(grouped_comm_reg, RMSE = sqrt(mean((actual_price - fcst_price)^2)))

# Calculate the Mean Absolute Percentage Error (MAPE) by index
mape_comm_reg <- summarise(grouped_comm_reg, MAPE = mean(abs((actual_price - fcst_price)/actual_price)) * 100)

# Print the resulting data frames
print(mae_comm_reg)
print(mse_comm_reg)
print(rmse_comm_reg)
print(mape_comm_reg)

# Heat map for RMSE
# create heatmap
ggplot(rmse_comm_reg, aes(x = Region, y = Commodity, fill = RMSE)) +
  geom_tile(color = "white") +
  scale_fill_gradient(low = "white", high = "blue") +
  labs(x = "Region", y = "Commodity", fill = "RMSE")

# Heat map for MAPE
# create heatmap
ggplot(mape_comm_reg, aes(x = Region, y = Commodity, fill = MAPE)) +
  geom_tile(color = "white") +
  scale_fill_gradient(low = "white", high = "orange") +
  labs(x = "Region", y = "Commodity", fill = "MAPE")

# Heat map for MAE
# create heatmap
ggplot(mae_comm_reg, aes(x = Region, y = Commodity, fill = MAE)) +
  geom_tile(color = "white") +
  scale_fill_gradient(low = "white", high = "red") +
  labs(x = "Region", y = "Commodity", fill = "MAE")


### COMMODITIY REGION YEAR METRICS ###
# Group the data by commodity_region_year
grouped_comm_reg_year <- group_by(data, Commodity, Region, Year)
grouped_comm_reg_year


# Calculate the Root Mean Squared Error (RMSE) by index
rmse_comm_reg_year <- summarise(grouped_comm_reg_year, RMSE = sqrt(mean((actual_price - fcst_price)^2)))


# Print the resulting data frames
print(rmse_comm_reg_year)

# Heat map for RMSE 2019
# create heatmap
rmse_comm_reg_2019 <- rmse_comm_reg_year %>%
  filter(Year == 2019)

ggplot(rmse_comm_reg_2019, aes(x = Region, y = Commodity, fill = RMSE)) +
  geom_tile(color = "white") +
  scale_fill_gradient(low = "white", high = "blue") +
  labs(x = "Region", y = "Commodity", fill = "RMSE") + 
  ggtitle("2019") + 
  theme(plot.title = element_text(hjust = 0.5))

# Heat map for RMSE 2020
# create heatmap
rmse_comm_reg_2020 <- rmse_comm_reg_year %>%
  filter(Year == 2020)

ggplot(rmse_comm_reg_2020, aes(x = Region, y = Commodity, fill = RMSE)) +
  geom_tile(color = "white") +
  scale_fill_gradient(low = "white", high = "blue") +
  labs(x = "Region", y = "Commodity", fill = "RMSE") + 
  ggtitle("2020") + 
  theme(plot.title = element_text(hjust = 0.5))

# Heat map for RMSE 2021
# create heatmap
rmse_comm_reg_2021 <- rmse_comm_reg_year %>%
  filter(Year == 2021)

ggplot(rmse_comm_reg_2021, aes(x = Region, y = Commodity, fill = RMSE)) +
  geom_tile(color = "white") +
  scale_fill_gradient(low = "white", high = "blue") +
  labs(x = "Region", y = "Commodity", fill = "RMSE") + 
  ggtitle("2021") + 
  theme(plot.title = element_text(hjust = 0.5))

# Heat map for RMSE 2022
# create heatmap
rmse_comm_reg_2022 <- rmse_comm_reg_year %>%
  filter(Year == 2022)

ggplot(rmse_comm_reg_2022, aes(x = Region, y = Commodity, fill = RMSE)) +
  geom_tile(color = "white") +
  scale_fill_gradient(low = "white", high = "blue") +
  labs(x = "Region", y = "Commodity", fill = "RMSE") + 
  ggtitle("2022") + 
  theme(plot.title = element_text(hjust = 0.5))

# Heat map for MAPE
# create heatmap
ggplot(mape_comm_reg, aes(x = Region, y = Commodity, fill = MAPE)) +
  geom_tile(color = "white") +
  scale_fill_gradient(low = "white", high = "orange") +
  labs(x = "Region", y = "Commodity", fill = "MAPE")

# Heat map for MAE
# create heatmap
ggplot(mae_comm_reg, aes(x = Region, y = Commodity, fill = MAE)) +
  geom_tile(color = "white") +
  scale_fill_gradient(low = "white", high = "red") +
  labs(x = "Region", y = "Commodity", fill = "MAE")


# visualize error metrics over time

data <- data %>%
  mutate(fcst_error = actual_price - fcst_price)

data <- data %>%
  rowwise() %>%
  mutate(mape = abs((actual_price - fcst_price) / actual_price) * 100,
         mae = abs(actual_price - fcst_price),
         mse = mean((actual_price - fcst_price)^2),
         rmse = sqrt(mean((actual_price - fcst_price)^2)))


head(data)


comm <- "Chlorine"
reg <- "APMKTS"
df_comm_reg <- data %>%
  filter(Commodity == comm & Region == reg)

df_agg <- df_comm_reg %>%
  group_by(Year, Month) %>%
  summarize(avg_MAPE = mean(mape))

# grouped by year
# ggplot(df_agg, aes(x = Month, y = avg_MAPE, group = Year, color = as.factor(Year))) +
#   geom_line() +
#   scale_x_continuous(breaks = 1:12, labels = month.name[1:12]) +
#   labs(title = paste0("Forecast Accuracy by Month - ", comm, " ", reg),
#        x = "Month",
#        y = "MAPE") +
#   theme_minimal()

# continuous
# ggplot(df_comm_reg, aes(x = year_month, y = fcst_error)) +
#   geom_line() +
#   labs(x = "Year Month", y = "Forecast Error")

#plot(df_comm_reg$fcst_error, type="l")

# grouped by scenario
df_agg_scenario <- df_comm_reg %>%
  group_by(Scenario) %>%
  summarize(avg_error = mean(fcst_error))

df_agg_scenario <- df_agg_scenario %>% 
  rowwise() %>%
  mutate(scenario_flipped = str_c(str_split(Scenario, "_")[[1]][2], "_", 
                         str_split(Scenario, "_")[[1]][1]))

ggplot(df_agg_scenario, aes(x = scenario_flipped, y = avg_error)) +
  geom_bar(stat = "identity") + 
  labs(x = "Fcst Scenario", y = "Avg Error") +
  theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust=1)) + 
  ggtitle(paste("Average Forecast Error by Scenario -", comm, reg))


# confidence intervals
# example for text: Caustic NA as of F07_2022

comm <- "Sodium Hydroxide"
reg <- "NA"
df_comm_reg_f07_2022 <- data %>%
  filter(Commodity == comm & Region == reg & 
           Scenario != "F08_2022" &
           Scenario != "F09_2022" &
           Scenario != "F10_2022" &
           Scenario != "F11_2022" &
           Scenario != "F12_2022")

ex_res <- df_comm_reg_f07_2022 %>%
  filter(Scenario != "F07_2022")
  
ex_res_mean <- mean(ex_res$fcst_error)  
ex_res_sd <- sd(ex_res$fcst_error)

ex_fcst <- df_comm_reg_f07_2022 %>%
  filter(Scenario == "F07_2022")

lower_ci <- ex_fcst$fcst_price - 1.96 * ex_res_sd
upper_ci <- ex_fcst$fcst_price + 1.96 * ex_res_sd

actuals_2022 <- actuals %>%
  filter(commodity == "SODIUM HYDROXIDE" & region == "NA" & year == 2022)

ex_for_graphing <- data.frame(
  date = seq(as.Date("2022-01-01"), as.Date("2022-12-01"), by = "month"),
  actual_price = c(actuals_2022$price[1],
                   actuals_2022$price[2],
                   actuals_2022$price[3],
                   actuals_2022$price[4],
                   actuals_2022$price[5],
                   actuals_2022$price[6],
                   actuals_2022$price[7],
                   actuals_2022$price[8],
                   actuals_2022$price[9],
                   actuals_2022$price[10],
                   actuals_2022$price[11],
                   actuals_2022$price[12]
  ),
  forecast_price = c(NA, NA, NA, NA, NA, NA, 
                     ex_fcst$fcst_price[1],
                     ex_fcst$fcst_price[2],
                     ex_fcst$fcst_price[3],
                     ex_fcst$fcst_price[4],
                     ex_fcst$fcst_price[5],
                     ex_fcst$fcst_price[6]
                     ),
  lower_95 = c(NA, NA, NA, NA, NA, NA,
               lower_ci[1],
               lower_ci[2],
               lower_ci[3],
               lower_ci[4],
               lower_ci[5],
               lower_ci[6]
               ),
  upper_95 = c(NA, NA, NA, NA, NA, NA,
               upper_ci[1],
               upper_ci[2],
               upper_ci[3],
               upper_ci[4],
               upper_ci[5],
               upper_ci[6]
  )
)

# line chart example
ggplot(ex_for_graphing, aes(x = date, y = actual_price)) +
  geom_line(color = "blue") +
  geom_line(aes(y = forecast_price), color = "red") +
  geom_ribbon(aes(ymin = lower_95, ymax = upper_95), fill = "gray", alpha = 0.3) +
  labs(title = "Sodium Hydroxide North America",
       x = "Date",
       y = "Price") +
  scale_x_date(date_labels = "%b %Y")

# get the forecast error from the previous forecast
ex_prev_fcst <- df_comm_reg_f07_2022 %>%
  filter(Scenario == "F06_2022")

# calc scaling factor
scaling_factor = 1 + (mean(ex_prev_fcst$fcst_error) - ex_res_mean) / ex_res_sd

# add adjusted forecast to data frame
ex_for_graphing$forecast_price_w_scale_factor = ex_for_graphing$forecast_price * scaling_factor

# plot with new line
ggplot(ex_for_graphing, aes(x = date, y = actual_price)) +
  geom_line(color = "blue") +
  geom_line(aes(y = forecast_price), color = "red") +
  geom_line(aes(y = forecast_price_w_scale_factor), color = "purple", linetype = "dotted") +
  geom_ribbon(aes(ymin = lower_95, ymax = upper_95), fill = "gray", alpha = 0.3) +
  labs(title = "Sodium Hydroxide North America",
       x = "Date",
       y = "Price") +
  scale_x_date(date_labels = "%b %Y")


# residual plot
# ggplot(df_agg_scenario, aes(x = forecast, y = residuals)) +
#   geom_point() +
#   geom_hline(yintercept = 0, color = "red") +
#   ggtitle("Residual Plot") +
#   xlab("Forecast") +
#   ylab("Residual")


##### Forecast Models #####

head(actuals)

acomm <- 'SODIUM HYDROXIDE'
areg <- 'NA'

# Load data and split into training and test sets

train <- actuals %>%
  filter(year < 2022 & 
           commodity == acomm & 
           region == areg)
  
test <- actuals %>%
  filter(year == 2022) %>%
  filter(commodity == acomm) %>%
  filter(region == areg)

# Fit an ARIMA model to the training set
model <- auto.arima(train$price)

# Generate forecasts for the test set
forecasts <- forecast(model, h=length(test$price))

# Calculate forecast errors
errors <- test$price - forecasts$mean

# Calculate mean and standard deviation of forecast errors
error_mean <- mean(errors)
error_sd <- sd(errors)

# Generate confidence intervals for forecasts
conf_int <- data.frame(
  lower = forecasts$mean - 1.96 * error_sd,
  upper = forecasts$mean + 1.96 * error_sd
)

# Combine forecasts and confidence intervals into a table
forecast_table <- cbind(test$price, forecasts$mean, conf_int)

# Plot the forecasts and confidence intervals
plot(forecasts, main="Forecast with Confidence Intervals")
lines(conf_int$lower, col="red", lty=2)
lines(conf_int$upper, col="red", lty=2)

# Calculate mean absolute error (MAE)
mae <- mean(abs(errors))

# Calculate mean absolute percentage error (MAPE)
mape <- mean(abs(errors / test$price)) * 100

# Calculate root mean squared error (RMSE)
rmse <- sqrt(mean(errors^2))

# testing fitting an ARIMA for every month
# to compare performance to CMA forecasts

# Specify the length of the forecast horizon
forecast_horizon <- 6

# Initialize an empty vector to store RMSE values
rmse_vals <- numeric(length = nrow(data) - forecast_horizon)

# Loop over the months, fitting ARIMA models and generating forecasts
for (i in seq(forecast_horizon + 1, nrow(data))) {
  # Extract the data up to the previous month
  train_data <- data[seq(1, i - forecast_horizon - 1), ]
  
  # Fit an ARIMA model to the previous month's data
  arima_model <- auto.arima(train_data$actual_price)
  
  # Generate a forecast for the next 6 months
  forecast <- forecast(arima_model, h = forecast_horizon)
  
  # Extract the actual prices for the next 6 months
  actual <- data$actual_price[seq(i, i + forecast_horizon - 1)]
  
  # Calculate RMSE for the forecast
  rmse_vals[i - forecast_horizon] <- sqrt(mean((forecast$mean - actual)^2))
}

# Calculate the average RMSE across all forecasts
mean_rmse <- mean(rmse_vals)



##### Model w Public Indicators #####
# Load your dataset into R
actuals_w_indicators <- read_excel("/Users/SamBarr7/OneDrive - UW-Eau Claire/Capstone/Commodity Actuals w Indicators.xlsx", sheet = "Export")

# Check the structure of your dataset
str(actuals_w_indicators)

# Split your dataset into training and testing sets
set.seed(123)
split <- sample.split(actuals_w_indicators$price, SplitRatio = 0.7)
train <- subset(actuals_w_indicators, split == TRUE)
test <- subset(actuals_w_indicators, split == FALSE)

# Fit a multiple linear regression model with your macroeconomic indicators
model <- lm(price ~ oil_price + scpi + ppi + sp_real + ip + unemployment, data = actuals_w_indicators)

# Check the summary of the model
summary(model)

# Make predictions on your testing set
predictions <- predict(model, newdata = test)

# Evaluate the performance of the model
rmse <- rmse(test$price, predictions)
mae <- mae(test$price, predictions)
r2 <- summary(model)$r.squared

# Print the performance metrics
print(paste0("RMSE: ", rmse))
print(paste0("MAE: ", mae))
print(paste0("R-squared: ", r2))

# run the regression for each commodity/region
# Fit segmented regression model for each item/region combination
models <- list()
for (commodity in unique(actuals_w_indicators$commodity)) {
  for (region in unique(actuals_w_indicators$region)) {
    subset_data <- actuals_w_indicators[actuals_w_indicators$commodity == commodity & actuals_w_indicators$region == region, ]
    model <- lm(price ~ oil_price + scpi + ppi + sp_real + ip + unemployment, data = subset_data)
    models[[paste0(commodity, "_", region)]] <- model
  }
}

# Print segmented regression results for each item/region combination
for (model_name in names(models)) {
  cat(model_name, ":\n")
  print(summary(models[[model_name]]))
}

# Bar plot of R-squared values for each item/region combination
r2_values <- sapply(models, function(model) summary(model)$r.squared)
sorted_r2_values <- sort(r2_values, decreasing = TRUE)
barplot(sorted_r2_values, 
        main = "R-squared Values for Regression Models", 
        ylab = "R-squared",
        las = 2,
        cex.names = 0.55)
par(mar = c(8, 4, 4, 2) + 0.1)

# plotting oil price to Chlorine NA
subset_data_for_scatter <- actuals_w_indicators[actuals_w_indicators$commodity == "CHLORINE" & actuals_w_indicators$region == "NA", ]
# Create a scatter plot of the actual data points
plot_data <- ggplot(subset_data_for_scatter, aes(x = oil_price, y = price)) + 
  geom_point() + 
  labs(x = "Year_Month", y = "Price", title = "TBD")

# Overlay the regression line
plot_data + geom_smooth(method = "lm", se = FALSE, formula = y ~ x, color = "red") + 
  ggtitle(paste("Regression line for item/region:", unique(subset_data_for_scatter$commodity), "-", unique(subset_data_for_scatter$region)))

### test a GLM model now
glm_model <- glm(price ~ oil_price + scpi + ppi + sp_real + ip + unemployment, data = actuals_w_indicators, family = Gamma(link = "log"))
summary(glm_model)

glm_models <- list()
for (commodity in unique(actuals_w_indicators$commodity)) {
  for (region in unique(actuals_w_indicators$region)) {
    subset_data <- actuals_w_indicators[actuals_w_indicators$commodity == commodity & actuals_w_indicators$region == region, ]
    glm_model <- glm(price ~ oil_price + scpi + ppi + sp_real + ip + unemployment, data = subset_data, family = Gamma(link = "log"))
    glm_models[[paste0(commodity, "_", region)]] <- glm_model
  }
}

for (glm_model_name in names(glm_models)) {
  cat(glm_model_name, ":\n")
  print(summary(glm_models[[glm_model_name]]))
}

# Bar plot of AIC values for each item/region combination
glm_aic_values <- sapply(glm_models, function(glm_model) summary(glm_model)$aic)
sorted_glm_aic_values <- sort(glm_aic_values, decreasing = FALSE)
barplot(sorted_glm_aic_values, 
        main = "AIC Values for Generalized Linear Models", 
        ylab = "AIC",
        las = 2,
        cex.names = 0.55)
par(mar = c(8, 4, 4, 2) + 0.1)
