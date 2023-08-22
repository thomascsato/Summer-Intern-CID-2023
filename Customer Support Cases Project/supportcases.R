library(readxl)
library(tidyverse)

# 6/5/2023 ----------------------------------------------------------------

# Exploration

SupportCases <- read_excel("All Support Cases 6-2-2023 Support Project.xlsx")
View(SupportCases)

str(SupportCases)
unique(SupportCases$`Case Age`)
unique(SupportCases$Product)
unique(SupportCases$Severity)
table(SupportCases$Severity)
unique(SupportCases$Status)
table(SupportCases$Status)
table(SupportCases$Resolution)

# Preliminary Graphs

ggplot(SupportCases, aes(fct_rev(fct_infreq(Product)))) +
  geom_bar(fill = "blue", color = "black") +
  coord_flip() +
  labs(x = "Product", y = "Number of Cases", title = "Number of Cases per Product") +
  theme_bw() +
  geom_text(aes(label = ..count..), stat = "count", hjust = -0.15) +
  scale_y_continuous(limit = c(0, 1300))

SupportCases %>%
  filter(!is.na(Product)) %>%
  ggplot(aes(fct_rev(fct_infreq(Product)))) +
    geom_bar(fill = "blue", color = "black") +
    coord_flip() +
    labs(x = "Product", y = "Number of Cases", title = "Number of Cases per Product") +
    theme_bw() +
    geom_text(aes(label = ..count..), stat = "count", hjust = -0.15) +
    scale_y_continuous(limit = c(0, 300))

SupportCases <- SupportCases %>%
  mutate(TimeUntilResolution = as.numeric(difftime(`Resolution Date`, `Created On`), units = "days")) %>%
  select(TimeUntilResolution, everything())

summary(SupportCases$TimeUntilResolution)

resolutiontime <- SupportCases %>%
  group_by(Product) %>%
  summarize(avgtime = mean(TimeUntilResolution, na.rm = TRUE))

View(resolutiontime)

View(SupportCases %>% filter(!is.na(TimeUntilResolution)))
View(SupportCases %>% filter(Product == "ModelBuilder"))

resolutiontime <- na.omit(resolutiontime)
ggplot(resolutiontime, aes(fct_reorder(Product, avgtime), avgtime)) +
  geom_col(fill = "darkgoldenrod", color = "black") +
  coord_flip() +
  labs(x = "Product", y = "Average Time Until Resolution (Days)", title = "Average Resolution Time by Product") +
  theme_bw()


lessthanzero <- SupportCases %>% filter(TimeUntilResolution <= 0)
View(lessthanzero)
nrow(lessthanzero)
# Not sure what to do with the 99 observations that have a resolution time that
# is less than zero.

resolutiontimezero <- SupportCases %>%
  filter(TimeUntilResolution > 0) %>%
  group_by(Product) %>%
  summarize(avgtime = mean(TimeUntilResolution, na.rm = TRUE))
resolutiontimezero <- na.omit(resolutiontimezero)

# It's pretty much the same graph though, the zeroes don't affect the mean all that much.
ggplot(resolutiontimezero, aes(fct_reorder(Product, avgtime), avgtime)) +
  geom_col(fill = "darkgoldenrod1", color = "black") +
  coord_flip() +
  labs(x = "Product", y = "Average Time Until Resolution (Days)", title = "Average Resolution Time by Product", caption = "Filtered for time differences < 0") +
  theme_bw()

# 6/6/2023

# Filling in product data
caselist <- SupportCases$`Case Title`
productcaseids <- str_match(caselist, "\\d{3}") # Matches first three digits in string
sum(is.na(productcaseids))

productdf <- tibble(unique(SupportCases$Product))
View(productdf)
productdf <- productdf %>%
  mutate(num = str_match(`unique(SupportCases$Product)`, "\\d{3}"))

SCTEST <- SupportCases
SCTEST %>%
  mutate(caseids = productcaseids) %>%
  left_join(productdf, by = c("caseids" = "num")) %>%
  select(`Case Title`, caseids, everything())

productdf <- productdf[-c(19, 21, 29, 30, 1, 22, 32, 23), ]
productdf[6, 2] <- "751 Avocado"
productdf[20, 2] <- "751 Melon"
productdf[11, 2] <- "751 Kiwi"
productdf[19, 2] <- "751 Mango"
productdf[21, 2] <- "901B"
productdf[22, 2] <- "901R"
productdf[23, 2] <- "901D"
productdf[7, 2] <- "901S"
colnames(productdf)[1] <- "Product"

# ************
SupportCases <- SupportCases %>%
  mutate(caseids = productcaseids) %>%
  left_join(productdf) %>%
  mutate(caseids = ifelse(is.na(caseids), num, caseids)) %>%
  left_join(productdf, by = c("caseids" = "num")) %>%
  select(-Product.x, -num, Product.y) %>%
  select(caseids, Product.y, everything())

sum(is.na(SupportCases$Product.y))
View(SupportCases[which(is.na(SupportCases$Product.y)),])
# Most of the missing product values are RMAs or web form submittals which are not useful.

RMAflag <- apply(SupportCases, 1, function(row){
  any(grepl("RMA", row))
})

SupportCases <- SupportCases %>%
  mutate(RMA = RMAflag) %>%
  select(RMA, everything())

# 699 RMA Cases in the dataset.
sum(SupportCases$RMA)

SupportCases <- SupportCases %>%
  mutate(caseids = as.vector(caseids))

# Exporting Data, filtering for NAs in description
SupportCasesDescriptions <- filter(SupportCases, !is.na(Description)) %>%
  select(Description, `Case Title`, Product.y, everything())

write_csv(SupportCasesDescriptions, "SupportCasesDescriptions.csv")

# Unnest tokens
descriptionstokens <- unnest_tokens(SupportCasesDescriptions, word, Description)

mostcommonwords <- descriptionstokens %>%
  group_by(word) %>%
  summarize(n = n()) %>%
  arrange(desc(n))

View(mostcommonwords)

productmostcommonwords <- descriptionstokens %>%
  group_by(Product.y, word) %>%
  summarize(n = n()) %>%
  arrange(desc(n)) %>%
  arrange(Product.y)

product_words <- function(product) {
  
  prodmostcommon <- productmostcommonwords %>%
    filter(Product.y == product)
  
  View(prodmostcommon)
  
}

prodlist <- unique(productmostcommonwords$Product.y)
product_words(prodlist[10])

# Want most common word for every description.
test1 <- filter(descriptionstokens, `Case Number` == "CAS-02136-H5F2X6" | `Case Number` == "CAS-06514-G0Y1L6")

test1 %>%
  group_by(`Case Number`, word) %>%
  summarize(n = n()) %>%
  group_by(`Case Number`) %>%
  filter(n == max(n))

descriptionmostcommon <- descriptionstokens %>%
  group_by(`Case Number`, word) %>%
  summarize(n = n()) %>%
  group_by(`Case Number`) %>%
  filter(n == max(n)) %>% # Multiple rows for each case
  group_by(`Case Number`) %>%
  mutate(word = paste0(word, collapse = ", ")) %>%
  distinct()

filter(descriptionmostcommon, `Case Number` == "CAS-01871-F0Q3N1")

SupportCasesDescriptions <- left_join(SupportCasesDescriptions, descriptionmostcommon) %>%
  select(Description, word, n, everything())

write_csv(SupportCasesDescriptions, "SupportCasesDescriptions2.csv")

# 6/13/2023

# Creating a dataframe of the missing values in the products
SupportCasesMissing <- SupportCases %>%
  filter(is.na(Product.y)) %>%
  select(Description, everything())

write_csv(SupportCasesMissing, "SupportCasesMissing.csv")

# Data frame of time resolutions that were less than 0
SupportCasesTimeLessThanZero <- SupportCases %>% filter(TimeUntilResolution <= 0) %>%
  select(`Created On`, `Resolution Date`, everything())

write_csv(SupportCasesTimeLessThanZero, "SupportCasesTimeLessThanZero.csv")

# Data frame of RMAs.
SupportCasesRMAs <- SupportCases %>%
  filter(RMA == TRUE) %>%
  select(Description, everything())

write_csv(SupportCasesRMAs, "SupportCasesRMAs.csv")

# 6/14/2023

# Filtering out NAs in resolution date and RMAs
# Theoretically, data frame of cases that were resolved, and were not resolved with RMA.

SupportCasesResolvedWithoutRMA <- SupportCases %>%
  filter(RMA == FALSE & Status == "Resolved")

nrow(SupportCasesResolvedWithoutRMA)
nrow(SupportCases %>% filter(RMA == FALSE & !is.na(`Resolution Date`)))

write_csv(SupportCasesResolvedWithoutRMA, "ResolvedWithoutRMA.csv")

# 6/15/2023

# Fun random data exploration

SupportCases <- SupportCases %>%
  mutate(nwordsdescription = sapply(str_split(Description, " "), FUN = length)) %>%
  select(nwordsdescription, everything())

# No relationship at all
ggplot(SupportCases, aes(nwordsdescription, TimeUntilResolution)) +
  geom_point()
