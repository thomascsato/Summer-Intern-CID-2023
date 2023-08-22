library(readxl)
library(tidyverse)
library(ggrepel)

Food <- read_excel("Market Analysis Long Range Planning.xlsx", sheet = 1)
Plant <- read_excel("Market Analysis Long Range Planning.xlsx", sheet = 2)
Spectroscopy <- read_excel("Market Analysis Long Range Planning.xlsx", sheet = 3)
Gas <- read_excel("Market Analysis Long Range Planning.xlsx", sheet = 4)
Image <- read_excel("Market Analysis Long Range Planning.xlsx", sheet = 5)

Food2 <- Food %>% 
  mutate(percentage = `TAM (USD)` / sum(`TAM (USD)`),
         csum = rev(cumsum(rev(`TAM (USD)`))), 
         pos = `TAM (USD)`/2 + lead(csum, 1),
         pos = if_else(is.na(pos), `TAM (USD)`/2, pos))

ggplot(Food, aes(x = "", y = `TAM (USD)`, fill = `Food Analysis Markets`)) +
  geom_col() +
  coord_polar(theta = "y") +
  theme_void() +
  labs(title = "Share of TAM for Food Analysis Markets") + 
  geom_label_repel(data = Food2,
                   aes(y = pos, label = paste0(round(percentage), "%")),
                   size = 4.5, nudge_x = 1, show.legend = FALSE)


piepercent <- round(100*Food$`TAM (USD)`/sum(Food$`TAM (USD)`), 1)

par(mar=c(2,2,2,20))
pie(Food$`TAM (USD)`, labels = piepercent, main = "City pie chart", col = rainbow(length(Food$`TAM (USD)`)))
legend(x = 0.1, y = 1, Food$`Food Analysis Markets`, cex = 0.8,
       fill = rainbow(length(Food$`TAM (USD)`)))




Food <- Food %>%
  mutate(`TAM (Billion USD)` = `TAM (USD)` / 1000000000,
         CAGR = CAGR * 100)

Plant <- Plant %>%
  mutate(`TAM (Billion USD)` = `...3` / 1000000000,
         CAGR = CAGR * 100)

Gas <- Gas %>%
  mutate(`TAM (Billion USD)` = `...3` / 1000000000,
         CAGR = CAGR * 100)

Spectroscopy <- Spectroscopy %>%
  mutate(`TAM (Billion USD)` = `...3` / 1000000000,
         CAGR = CAGR * 100)

Image <- Image %>%
  mutate(`TAM (Billion USD)` = `...3` / 1000000000,
         CAGR = CAGR * 100)

ggplot(Food, aes(fct_rev(fct_reorder(`Food Analysis Markets`, `TAM (Billion USD)`)), `TAM (Billion USD)`)) +
  geom_col(fill = "blue4") +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1)) +
  geom_text(aes(label = round(`TAM (Billion USD)`, 2)), vjust = -0.5) +
  labs(x = "Food Analysis Markets", title = "Food Analysis Markets TAM (Billion USD)") +
  theme(axis.text = element_text(size = 10),
        axis.title = element_text(size = 14),
        plot.title = element_text(size = 16))
  

ggplot(Food, aes(fct_reorder(`Food Analysis Markets`, `TAM (Billion USD)`), `TAM (Billion USD)`)) +
  geom_col(fill = "blue4") +
  theme_minimal() +
  geom_text(aes(label = sprintf("%.2f", `TAM (Billion USD)`), vjust = 0.2, hjust = -0.2)) +
  labs(x = "Food Analysis Markets", title = "Food Analysis Markets TAM (Billion USD)") +
  coord_flip() +
  scale_y_continuous(limits = c(0, max(Food$`TAM (Billion USD)`, na.rm = T) + ceiling(max(Food$`TAM (Billion USD)`, na.rm = T) * 0.05)))

ggplot(Plant, aes(fct_reorder(`Plant Analysis Markets`, `TAM (Billion USD)`), `TAM (Billion USD)`)) +
  geom_col(fill = "green4") +
  theme_minimal() +
  geom_text(aes(label = sprintf("%.2f", `TAM (Billion USD)`), vjust = 0.2, hjust = -0.2)) +
  labs(x = "Plant Analysis Markets", title = "Plant Analysis Markets TAM (Billion USD)") +
  coord_flip() +
  scale_y_continuous(limits = c(0, max(Plant$`TAM (Billion USD)`, na.rm = T) + ceiling(max(Plant$`TAM (Billion USD)`, na.rm = T) * 0.1)))

ggplot(Gas, aes(fct_reorder(`Gas Analysis Markets`, `TAM (Billion USD)`), `TAM (Billion USD)`)) +
  geom_col(fill = "lightblue4") +
  theme_minimal() +
  geom_text(aes(label = sprintf("%.2f", `TAM (Billion USD)`), vjust = 0.3, hjust = -0.2)) +
  labs(x = "Gas Analysis Markets", title = "Gas Analysis Markets TAM (Billion USD)") +
  coord_flip() +
  scale_y_continuous(limits = c(0, max(Gas$`TAM (Billion USD)`, na.rm = T) + ceiling(max(Gas$`TAM (Billion USD)`, na.rm = T) * 0.1)))

ggplot(Spectroscopy, aes(fct_reorder(`Spectroscopy Markets`, `TAM (Billion USD)`), `TAM (Billion USD)`)) +
  geom_col(fill = "red4") +
  theme_minimal() +
  geom_text(aes(label = sprintf("%.2f", `TAM (Billion USD)`), vjust = 0.2, hjust = -0.2)) +
  labs(x = "Spectroscopy Markets", title = "Spectroscopy Markets TAM (Billion USD)") +
  coord_flip() +
  scale_y_continuous(limits = c(0, max(Spectroscopy$`TAM (Billion USD)`, na.rm = T) + ceiling(max(Spectroscopy$`TAM (Billion USD)`, na.rm = T) * 0.1)))

ggplot(Image, aes(fct_reorder(`Image Analysis Markets`, `TAM (Billion USD)`), `TAM (Billion USD)`)) +
  geom_col(fill = "yellow4") +
  theme_minimal() +
  geom_text(aes(label = sprintf("%.2f", `TAM (Billion USD)`), vjust = 0.3, hjust = -0.2)) +
  labs(x = "Image Analysis Markets", title = "Image Analysis Markets TAM (Billion USD)") +
  coord_flip() +
  scale_y_continuous(limits = c(0, max(Image$`TAM (Billion USD)`, na.rm = T) + ceiling(max(Image$`TAM (Billion USD)`, na.rm = T) * 0.1)))


# CAGR

ggplot(Food, aes(fct_reorder(`Food Analysis Markets`, CAGR), CAGR)) +
  geom_col(fill = "blue4") +
  theme_minimal() +
  geom_text(aes(label = paste(sprintf("%.2f", CAGR), "%"), vjust = 0.2, hjust = -0.2)) +
  labs(x = "Food Analysis Markets", title = "Food Analysis Markets CAGR") +
  coord_flip() +
  scale_y_continuous(limits = c(0, max(Food$CAGR, na.rm = T) + max(Food$CAGR, na.rm = T) * 0.1))

ggplot(Plant, aes(fct_reorder(`Plant Analysis Markets`, CAGR), CAGR)) +
  geom_col(fill = "green4") +
  theme_minimal() +
  geom_text(aes(label = paste(sprintf("%.2f", CAGR), "%"), vjust = 0.2, hjust = -0.2)) +
  labs(x = "Plant Analysis Markets", title = "Plant Analysis Markets CAGR") +
  coord_flip() +
  scale_y_continuous(limits = c(0, max(Plant$CAGR, na.rm = T) + ceiling(max(Plant$CAGR, na.rm = T) * 0.1)))

ggplot(Gas, aes(fct_reorder(`Gas Analysis Markets`, CAGR), CAGR)) +
  geom_col(fill = "lightblue4") +
  theme_minimal() +
  geom_text(aes(label = paste(sprintf("%.2f", CAGR), "%"), vjust = 0.3, hjust = -0.2)) +
  labs(x = "Gas Analysis Markets", title = "Gas Analysis Markets CAGR") +
  coord_flip() +
  scale_y_continuous(limits = c(0, max(Gas$CAGR, na.rm = T) + ceiling(max(Gas$CAGR, na.rm = T) * 0.1)))

ggplot(Spectroscopy, aes(fct_reorder(`Spectroscopy Markets`, CAGR), CAGR)) +
  geom_col(fill = "red4") +
  theme_minimal() +
  geom_text(aes(label = paste(sprintf("%.2f", CAGR), "%"), vjust = 0.2, hjust = -0.2)) +
  labs(x = "Spectroscopy Markets", title = "Spectroscopy Markets CAGR") +
  coord_flip() +
  scale_y_continuous(limits = c(0, max(Spectroscopy$CAGR, na.rm = T) + ceiling(max(Spectroscopy$CAGR, na.rm = T) * 0.1)))

ggplot(Image, aes(fct_reorder(`Image Analysis Markets`, CAGR), CAGR)) +
  geom_col(fill = "yellow4") +
  theme_minimal() +
  geom_text(aes(label = paste(sprintf("%.2f", CAGR), "%"), vjust = 0.3, hjust = -0.2)) +
  labs(x = "Image Analysis Markets", title = "Image Analysis Markets CAGR") +
  coord_flip() +
  scale_y_continuous(limits = c(0, max(Image$CAGR, na.rm = T) + ceiling(max(Image$CAGR, na.rm = T) * 0.1)))
