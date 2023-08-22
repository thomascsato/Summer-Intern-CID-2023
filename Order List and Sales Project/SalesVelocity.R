# This is the whole messy script that I used to explore the data and everything-- The Marketing Project.Rmd is better code to look at.





library(tidyverse)
library(readxl)
library(lubridate)

May23out <- read_excel("Orderlist 2023.xlsm", sheet = "May 23 Out",
                       skip = 1, col_types = "text")
May23in <- read_excel("Orderlist 2023.xlsm", sheet = "May 23 In",
                      col_types = "text")

monthslists <- excel_sheets("Orderlist 2023.xlsm")[c(9:10, 12:58, 60:92)]

April23out <- read_excel("Orderlist 2023.xlsm", sheet = "April 23 Out",
                         col_types = "text")
April23in <- read_excel("Orderlist 2023.xlsm", sheet = "April 23 In",
                        col_types = "text")

# Binding rows test
View(bind_rows(May23out, April23out))

# Now to figure out how to do this for all of the datasets
outs <- monthslists[seq(1, length(monthslists), 2)][-1]
ins <- monthslists[-seq(1, length(monthslists), 2)][-1]

colnames(April23in)[1] <- "Source"
colnames(May23in)[1] <- "Source"

View(bind_rows(May23in, April23in))

outdatasets <- May23out
for(i in seq(1:length(outs))) {
  if(i < 26) {
    outdataset <- read_excel("Orderlist 2023.xlsm", sheet = outs[i],
                            col_types = "text")
  } else {
    outdataset <- read_excel("Orderlist 2023.xlsm", sheet = outs[i],
                             col_types = "text", skip = 1)
  }
  outdatasets <- bind_rows(outdatasets, outdataset)
}

View(outdatasets)

# ATTENTION:  There may be a need for more data cleaning after this step.
# I suspect so because there are a lot of added columns that should not be there
# I suspect those are relics from the data formats from 2020 or thereabouts.
# NOTE 5/31/23 - I found one of the columns
View(filter(outdatasets, startsWith(`...32`, "950")))
# Column `...32` comes from Dec 2020 Out dataset.
# I think it is safe to say that the extra columns are not necessary

outshipping <- filter(outdatasets[, 1:18],
                      !is.na(Source) | Source %in% c("REPAIR", "Repair", "repair"))

indatasets <- May23in
for(i in seq(1:length(ins))) {
  indataset <- read_excel("Orderlist 2023.xlsm", sheet = ins[i],
                          col_types = "text")
  colnames(indataset)[1] <- "Source"
  indatasets <- bind_rows(indatasets, indataset)
}

inorders <- filter(indatasets[, 1:11], 
                   !is.na(Source) | Source %in% c("REPAIR", "Repair", "repair"))

# Filtering out the repair orders -- There are 428 repair orders from 2020 onward
outshipping <- filter(outshipping, !startsWith(`SO#`, "RMA"))


# Joining the Tables Together ---------------------------------------------

(orders <- left_join(outshipping, inorders, by = c(`SO#` = "SO")))

# This tells us that 81 of the orders here did not have matching SO#'s.
sum(is.na(orders$Source.y))


# For some reason, the dates got converted to numbers that are 70 years in the
# future and also one day later.  The one day off I suppose is probably due to time
# zones or leap years I hypothesize, but otherwise it has the right month and day.

# The first observation in the May23in dataset for the date ordered.
as_date(45047)


# Date and Time Series Analysis -------------------------------------------

orders$Ordered

orders$Ordered <- as_date(as.numeric(orders$Ordered) - 25568)
range(orders$Ordered, na.rm = T)
View(orders)


ordersdates <- filter(orders, !is.na(Ordered))
ordersdates <- filter(ordersdates, Ordered < "2023-06-01")

ordersdates$QTY <- ifelse(is.na(ordersdates$QTY), 1, ordersdates$QTY)

ggplot(ordersdates) +
  geom_bar(aes(Ordered))

ordersdatesqty <- ordersdates %>%
  group_by(Ordered) %>%
  summarize(sum_orders = sum(as.numeric(QTY), na.rm = TRUE))

ggplot(ordersdatesqty) +
  geom_line(aes(Date, sum_orders)) +
  labs(x = "Date (Day)", y = "Quantity of Orders",
       title = "Number of Orders from 2020-Present") +
  theme_bw()

ordersdatesqty <- ordersdatesqty %>%
  mutate(Day = day(Ordered), Month = month(Ordered), Year = year(Ordered))

ordersdatesqty1 <- ordersdatesqty %>%
  group_by(Month, Year) %>%
  summarize(sum_orders = sum(sum_orders, na.rm = TRUE)) %>%
  unite(col = "Date", c("Year", "Month"), sep = "-") %>%
  mutate(Date = as_date(paste(Date, "-01")))

ggplot(ordersdatesqty1) +
  geom_line(aes(Date, sum_orders), size = 1.2) +
  labs(x = "Date (Month)", y = "Quantity of Orders",
       title = "Number of Orders by Month") +
  theme_bw()





# Duplicates --------------------------------------------------------------

duplicates <- outshipping[which(duplicated(outshipping$`SO#`)), ]
SOs <- duplicates$`SO#`
View(filter(outshipping, `SO#` %in% SOs))

sum(duplicated(orders$`SO#`))
ordersduplicates <- orders[which(duplicated(orders$`SO#`)), ]
orderSOs <- ordersduplicates$`SO#`
View(filter(orders, `SO#` %in% orderSOs))




#6/1/2023


# Individual Products ------------------------------------------------------

# This is if I want to standardize the exact product that is being sold
# Instead, for now I want to categorize by company (CID, Felix, Interscan)

products <- c("F-920", "CI-202", "F-950", "CI-710", "AVO", "F-750", "710", "CI-203",
              "MANGO", "F-960", "CI-110", "CI-600", "CI-340", "F-900", "F-901", "CI-602",
              "CI-510", "F-940", "KIWI", "CI-301")
test <- ordersdates

# The issue is that if one specific product CONTAINS the substring ("F-950" for example), I want
# it to match that substring with the corresponding string in the products variable.
test$PRODUCTS <- ifelse(stri_detect_fixed(paste(products, collapse = "|"), test$PRODUCTS),
                        products[products == 123], # This line is the one that needs to be figured out.
                        "Miscellaneous")

#In order to visualize a time series based on product (or company), I need to standardize the text in the product columns, so that instead of "F-750 AVO LOANER PURCHASE" or "F-751 AVO LOANER\\r\\nF-750 LOANER", it would say "F-751 AVO".



# Company Names -----------------------------------------------------------

ordersdates
productlist <- ordersdates$`Accessories/Details`
# Test for the Regex \\d{3}
str_match(c("123", "vksjd234nskjvn456", "1 00pvmdps98 7651 23km098cdkj"), "\\d{3}")

productids <- str_match(productlist, "\\d{3}")
sum(is.na(productids))

ordersdates <- ordersdates %>% 
  mutate(ordersdates, product_IDs = productids) %>%
  select(product_IDs, everything())

unique(ordersdates$product_IDs)

# Potential problem:  There is an Interscan product with the num 900 and a Felix product with 900
cid <- c(202, 710, 602, 600, 203, 110, 340, 301)
felix <- c(920, 950, 750, 751, 960, 940, 901, 900)
interscan <- c(900)

# Converts to NA by default else, can specify .default argument
companycategories <- case_when(
  ordersdates$product_IDs %in% cid ~ "CID",
  ordersdates$product_IDs %in% felix & startsWith(ordersdates$PRODUCTS, "F") ~ "Felix",
  ordersdates$product_IDs %in% interscan ~ "Interscan",
  tolower(ordersdates$Source.x) == "interscan" ~ "Interscan"
)

ordersdates <- ordersdates %>%
  mutate(Company = companycategories) %>%
  select(Company, everything())

View(filter(ordersdates, is.na(Company)))


# 6/5/2023

ordersdates <- mutate(ordersdates, Year = year(Ordered)) %>% select(Ordered, Year, everything())

# Might be a good graph to have once more years of sales are integrated.
ggplot(ordersdates) +
  geom_bar(aes(Year))

ordersdates <- ordersdates %>%
  mutate(YD = as.yearmon(Ordered)) %>%
  select(YD, everything())

dates <- ordersdates %>%
  group_by(YD, Company) %>%
  summarize(n = n()) %>%
  filter(!is.na(Company))

ggplot(dates) +
  geom_line(aes(as.Date(YD), n, group = Company, color = Company)) +
  labs(x = "Date", y = "Number of Sales", title = "Number of Sales by Company") +
  theme_bw() +
  scale_x_date(breaks = as.Date(dates$YD)) +
  theme(axis.text.x = element_text(angle=45, hjust=1)) # Optional, might be better to stick with yearly labels
  
ggplot(dates) +
  geom_line(aes(YD, n, group = Company, color = Company)) +
  labs(x = "Date", y = "Number of Sales", title = "Number of Sales by Company") +
  theme_bw()

# Look at time of week now.
ordersdates <- mutate(ordersdates, wday = wday(Ordered, label = T, abbr = F)) %>%
  select(wday, everything())
ordersdates

ggplot(ordersdates) +
  geom_bar(aes(wday), fill = "lightblue", color = "black") +
  labs(x = "Day of the Week", y = "Number of Sales", title = "Number of Sales by Day Ordered") +
  theme_bw()

ggplot(ordersdates) +
  geom_bar(aes(wday, fill = Company),color = "black", position = "dodge") +
  labs(x = "Day of the Week", y = "Number of Sales", title = "Number of Sales by Day Ordered") +
  theme_bw()

weekdaycounts <- ordersdates %>%
  group_by(wday) %>%
  summarize(n = n())



dts <- ordersdates %>%
  group_by(wday, Company) %>%
  summarize(n = n())

ggplot(dts, aes(wday, n, fill = Company)) +
  geom_col(position = "dodge") +
  geom_text(aes(label = case_when(
    Company == "CID" ~ sprintf(n/sum(filter(dts, Company == "CID")$n), fmt = "%#.2f"),
    Company == "Felix" ~ sprintf(n/sum(filter(dts, Company == "Felix")$n), fmt = "%#.2f"),
    Company == "Interscan" ~ sprintf(n/sum(filter(dts, Company == "Interscan")$n), fmt = "%#.2f"),
    is.na(Company) ~ sprintf(n/sum(filter(dts, is.na(Company))$n), fmt = "%#.2f")
    )), 
    position = position_dodge(.9), 
    vjust = -0.5, 
    size = 3.5)

ggplot(ordersdates) +
  geom_bar(aes(wday, fill = Company),color = "black", position = "dodge") +
  labs(x = "Day of the Week", y = "Number of Sales", title = "Number of Sales by Day Ordered") +
  theme_bw()





# Briefly trying to find number of unique GasD customers for Scott
View(orderdates)

names(ordersdates)[6] <- "productID"
names(ordersdates)
filter(ordersdates, productID == 900)
GasD <- subset(ordersdates, productID == "900")
length(unique(GasD$`CUSTOMER'S NAME`))
# Creates column of each company corresponding to the number in the product_IDs column
productlist <- outdatasets$`Accessories/Details`

# The Regex "\\d{3}" matches the first three consecutive numbers of each string in productlist
productids <- str_match(productlist, "\\d{3}")
outdatasets <- outdatasets %>% 
  mutate(outdatasets, product_IDs = productids) %>%
  select(product_IDs, everything())

GasD <- filter(outdatasets, product_IDs == "900")
length(unique(GasD$`CUSTOMER'S NAME`))

# Looks like 51 unique customers


# 6/7/2023

# Need to see why there are only 9 observations in the April 2021 data
inordersordered <- as_date(as.numeric(inorders$Ordered) - 25569)
apr21indates <- inordersordered[year(inordersordered) == 2021 & month(inordersordered) == 4]
# There were 37 ordered in april of 21

datechangeinorders <- inorders %>%
  mutate(Ordered = as_date(as.numeric(inorders$Ordered) - 25569))
datechangeinorders
apr21ins <- filter(datechangeinorders, Ordered %in% apr21indates)
apr21ins$SO

apr21outs <- filter(outshipping, `SO#` %in% apr21ins$SO)
View(apr21outs)

datechangeoutshipping <- outshipping %>%
  mutate(Ordered = as_date(as.numeric(`PACK DATE`) - 25569))
apr21outs <- filter(datechangeoutshipping, Ordered %in% Ordered[year(Ordered) == 2021 & month(Ordered) == 2])
View(apr21outs)

# Redoing the original analysis from 2020-present day.


# REDOING 2020-present ANALYSIS -------------------------------------------
# A lot of this is going to be repeated code from above

excel_sheets("Orderlist 2023.xlsm")
monthslists <- excel_sheets("Orderlist 2023.xlsm")[c(12:93)]

outs <- monthslists[seq(1, length(monthslists), 2)][-1]
ins <- monthslists[-seq(1, length(monthslists), 2)][-1]

# Skips the first line for May 23 so that the column names are what they should be
May23out <- read_excel("Orderlist 2023.xlsm", sheet = "May 23 Out",
                       skip = 1, col_types = "text")
May23in <- read_excel("Orderlist 2023.xlsm", sheet = "May 23 In",
                      col_types = "text")

# Renaming the first column so that the format is standardized
colnames(May23in)[1] <- "Source"

outdatasets <- May23out

for(i in seq(1:length(outs))) {
  
  # From January 2020 - March 2021, the formatting was different, which is why we skip a line past the 26th name in the outs list.
  if(i < 26) {
    outdataset <- read_excel("Orderlist 2023.xlsm", sheet = outs[i],
                             col_types = "text")
  } else {
    outdataset <- read_excel("Orderlist 2023.xlsm", sheet = outs[i],
                             col_types = "text", skip = 1)
  }
  
  # Binds the rows to the outdatasets dataset
  outdatasets <- bind_rows(outdatasets, outdataset)
}

indatasets <- May23in

for(i in seq(1:length(ins))) {
  
  indataset <- read_excel("Orderlist 2023.xlsm", sheet = ins[i],
                          col_types = "text")
  
  # Renaming the first column of the dataset to match what it is in the "Out" datasets, and to be able to bind the rows together for that column
  colnames(indataset)[1] <- "Source"
  
  indatasets <- bind_rows(indatasets, indataset)
}

outshipping <- filter(outdatasets[, 1:18],
                      !is.na(Source) | !(Source %in% c("REPAIR", "Repair", "repair")))

# Filtering out repairs 
outshipping <- filter(outshipping, !startsWith(`SO#`, "RMA"))

inorders <- filter(indatasets[, 1:11], 
                   !is.na(Source) | !(Source %in% c("REPAIR", "Repair", "repair")))

# Getting rid of one leftover observation that starts with "RMA"
inorders <- filter(inorders, !startsWith(SO, "RMA"))

(orders <- left_join(outshipping, inorders, by = c(`SO#` = "SO")))
orders$Ordered <- as_date(as.numeric(orders$Ordered) - 25569)

ordersdates <- filter(orders, !is.na(Ordered))
ordersdates <- filter(ordersdates, Ordered < "2023-06-01")

numsales <- ordersdates %>%
  group_by(Ordered) %>%
  summarize(num_sales = n())

dates20to23 <- as.data.frame(seq(as.Date("2020-01-01"), as.Date("2023-05-31"), by="days"))
names(dates20to23) <- "Date"

numsales <- left_join(dates20to23, numsales, by = c("Date" = "Ordered")) %>%
  mutate(num_sales = replace_na(num_sales, 0),
         Day = day(Date),
         Month = month(Date),
         Year = year(Date))

monthly_numsales <- numsales %>%
  group_by(Year, Month) %>%
  summarize(num_sales = sum(num_sales))

# Creating a time series object for the sales per month
monthlysales <- ts(monthly_numsales$num_sales, start = 2020, frequency = 12)
monthlysales

# For 2020
datechangeinorders <- inorders %>%
  mutate(Ordered = as_date(as.numeric(inorders$Ordered) - 25569))
apr21outs <- filter(datechangeinorders, Ordered %in% Ordered[year(Ordered) == 2021 & month(Ordered) == 4])
apr21ins$SO # 39 unique SO#s

datechangeoutshipping <- outshipping %>%
  mutate(Ordered = as_date(as.numeric(`PACK DATE`) - 25569))
apr21outs <- filter(datechangeoutshipping, Ordered %in% Ordered[year(Ordered) == 2021 & month(Ordered) == 4])
apr21outs$`SO#`


# 6/8/2023 Still looking for the April 2021 discrepency

which(outdatasets$`PACK DATE` == "2021-04-20")
outdatasets <- outdatasets %>%
  mutate(`PACK DATE` = as_date(as.numeric(`PACK DATE`) - 25569))

outdatasets <- outdatasets %>%
  mutate(md = as.yearmon(`PACK DATE`))
filter(outdatasets, md == "May 2023")

# Now to try and get rid of duplicate sales.

View(orders)

View(orders[duplicated(orders$`SO#`), ])

sum(duplicated(orders))

# Duplicates again

duplicates <- outshipping[which(duplicated(outshipping$`SO#`)), ]
SOs <- duplicates$`SO#`
View(filter(outshipping, `SO#` %in% SOs))

sum(duplicated(orders$`SO#`))
ordersduplicates <- orders[duplicated(orders$`SO#`), ]
orderSOs <- ordersduplicates$`SO#`
allduplicates <- filter(orders, `SO#` %in% orderSOs) %>% arrange(`SO#`)


# 6/13/23

# Attempt at forecasting?

etsmonthlysales <- ets(monthlysales)
forecast(etsmonthlysales, h = 10)
plot(forecast(etsmonthlysales, h = 10))

# ARIMA model thinks that a constant point forecast is most accurate AKA there is no forecasting potential here
arima.model <- auto.arima(monthlysales)
forecast(arima.model,h=10)
plot(forecast(arima.model,h=10))

# 6/14/23

# Attempt at matching product numbers to names for specific products
# Using this because it has a good list of products
SupportCases <- read_excel("All Support Cases 6-2-2023 Support Project.xlsx")

productdf <- tibble(unique(SupportCases$Product))
productdf <- productdf %>%
  mutate(num = str_match(`unique(SupportCases$Product)`, "\\d{3}"))
productdf <- productdf[-c(19, 21, 29, 30, 1, 22, 32, 23, 33), ]
productdf[6, 2] <- "751 Avocado"
productdf[20, 2] <- "751 Melon"
productdf[11, 2] <- "751 Kiwi"
productdf[19, 2] <- "751 Mango"
productdf[21, 2] <- "901B"
productdf[22, 2] <- "901R"
productdf[23, 2] <- "901D"
productdf[7, 2] <- "901S"
colnames(productdf)[1] <- "Product"
productdf$num <- as.vector(productdf$num)


View(productdf)

# Using case_when again in order to change the specific products that have a common number
# i.e. 751 MAN, 751 AVO, 751 KIWI
companycategories <- case_when(
  ordersdates$product_IDs %in% cid ~ "CID",
  tolower(ordersdates$Source.x) == "interscan" ~ "Interscan",
  tolower(ordersdates$Source.y) == "interscan" ~ "Interscan",
  ordersdates$product_IDs %in% felix ~ "Felix"
)

productids4 <- case_when(
  grepl("avo", tolower(ordersdates$Product)) ~ "751 Avocado",
  grepl("man", tolower(ordersdates$Product)) ~ "751 Mango",
  grepl("kiwi", tolower(ordersdates$Product)) ~ "751 Kiwi",
  grepl("901B", ordersdates$Product) ~ "901B",
  grepl("901R", ordersdates$Product) ~ "901R",
  grepl("901D", ordersdates$Product) ~ "901D",
)

productids3 <- case_when(
  grepl("avo", tolower(ordersdates$PRODUCTS)) ~ "751 Avocado",
  grepl("man", tolower(ordersdates$PRODUCTS)) ~ "751 Mango",
  grepl("kiwi", tolower(ordersdates$PRODUCTS)) ~ "751 Kiwi",
  grepl("901B", ordersdates$PRODUCTS) ~ "901B",
  grepl("901R", ordersdates$PRODUCTS) ~ "901R",
  grepl("901D", ordersdates$PRODUCTS) ~ "901D",
)

orderstest <- ordersdates
orderstest$product_IDs <- ifelse(!is.na(productids4), productids4, orderstest$product_IDs)
orderstest$product_IDs <- ifelse(!is.na(productids3), productids3, orderstest$product_IDs)
View(filter(orderstest, product_IDs == "751"))

orderstest <- orderstest %>%
  left_join(productdf, by = c("product_IDs" = "num")) %>%
  select(Product.y, everything())
