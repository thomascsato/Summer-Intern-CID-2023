library(readxl)
library(tidyverse)

InterscanMay2023 <- read_excel("Interscan Sensor Orderlist 2023 0808.xlsx", sheet = "May OUT 2023")
InterscanJun2023 <- read_excel("Interscan Sensor Orderlist 2023 0808.xlsx", sheet = "June OUT 2023")
InterscanJul2023 <- read_excel("Interscan Sensor Orderlist 2023 0808.xlsx", sheet = "July OUT 2023")

InterscanJul2023$`SHIP DATE` <- as.character(InterscanJul2023$`SHIP DATE`)

Interscan <- bind_rows(InterscanMay2023, InterscanJun2023)
Interscan <- bind_rows(Interscan, InterscanJul2023)

interscanordersdates <- ordersdates %>% filter(Company == "Interscan")

unique(interscanordersdates$customer)

1 -pnorm(.5)

# Generate a sequence of x values
x <- seq(-5, 5, length=100)

# Calculate the PDF values for a standard normal distribution
pdf_values <- 1/sqrt(2*pi) - dnorm(x, mean=0, sd=1)  # mean=0, sd=1 for standard normal

# Create the plot
plot(x, pdf_values, type="l", lwd=2, col="blue", 
     xlab="X", ylab="Probability Density", main="Standard Normal Distribution PDF")

# height_at_center = norm.pdf(0, loc=0, scale=1) in scipy