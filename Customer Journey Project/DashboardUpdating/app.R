#
# This is a Shiny web application. You can run the application by clicking
# the 'Run App' button above.
#
# Find out more about building applications with Shiny here:
#
#    http://shiny.rstudio.com/
#

library(shiny)
library(tidyverse)
library(tidytext)
library(readxl)
library(lubridate)
library(zoo)

ui <- fluidPage(
  
  fileInput("file1", "Upload Orderlist"),
  fileInput("file2", "Upload Support Cases"),
  fileInput("file6", "Upload Published Papers"),
  fileInput("file7", "Upload Published Papers Terms"),
  
  fluidRow(
    column(1,
      actionButton("modifyBtn", "Modify Orderlist"),
      actionButton("modifyBtn5", "Modify Published Papers"),
      downloadButton("downloadOrderlist", "Download Modified Orderlist"),
      downloadButton("downloadPapers", "Download Modified Published Papers"),
    )
  ),
  
  # UI progress bar
  verbatimTextOutput("output"),
  uiOutput("progressBar1"),
  uiOutput("progressBar5")
  
)

server <- function(input, output) {
  
  # Allows user to upload larger files
  options(shiny.maxRequestSize = 10*1024^2)
  
  # Rendering of the progress bar
  # setProgress indicates when and how much I want the progress bar to fill
  output$progressBar1 <- renderUI({
    if (input$modifyBtn > 0) {
      progressBar <- withProgress(message = "Processing files...",
                                  detail = "Please wait...",
                                  value = 0, {
                                    setProgress(0)
                                  })
      progressBar
    }
  })
  
  output$progressBar5 <- renderUI({
    if (input$modifyBtn5 > 0) {
      progressBar <- withProgress(message = "Processing files...",
                                  detail = "Please wait...",
                                  value = 0, {
                                    setProgress(0)
                                  })
      progressBar
    }
  })
  
  observeEvent(input$modifyBtn, {
    
    # Saving the input files to variables
    Orderlist <- input$file1
    Support <- input$file2
    
    # Will show a message if all files are not uploaded
    if(is.null(Orderlist) || is.null(Support)) {
      
      output$output <- renderPrint({
        "Please upload Orderlist and Support Cases."
      })
      return()
      
    }
    
    # Indicating that I want the progress bar to show up when this code begins (AKA the processing begins)
    withProgress(session = shiny::getDefaultReactiveDomain(),
                 message = "Processing files...",
                 detail = "Please wait...",
                 value = 0, {
                   
      # Where all of the data processing and cleaning begins.
      # Where there is a .xlsx file, it is replaced with the corresponding input$datapath.
      
      monthslists <- excel_sheets(Orderlist$datapath)
      
      outs <- monthslists[grepl("out", tolower(monthslists))]
      ins <- monthslists[grepl("in", tolower(monthslists))]
      
      # First sheet in Orderlist, should be most recent OUT.
      outdatasets <- read_excel(Orderlist$datapath, sheet = outs[1],
                               skip = 1, col_types = "text")
      
      setProgress(0.1)
      
      for(i in seq(1:(length(outs)))) {
       
        outdataset <- read_excel(Orderlist$datapath, sheet = outs[i],
                                col_types = "text", skip = 1)
        
        outdatasets <- bind_rows(outdatasets, outdataset)
        
        setProgress(0.1 + i/(length(outs))/100)
       
      }
      
      outdatasets <- outdatasets %>%
        mutate(`Accessories/Details` = ifelse(is.na(`Accessories/Details`), `PRODUCTS DESCRIPTION`, `Accessories/Details`),
              `Accessories/Details` = ifelse(is.na(`Accessories/Details`), PRODUCTS, `Accessories/Details`),
              `SO#` = ifelse(is.na(`SO#`), `INV / SO #`, `SO#`),
              `SO#` = ifelse(is.na(`SO#`), `SO #`, `SO#`)) %>%
        select(`Accessories/Details`, everything())
      
      indatasets <- read_excel(Orderlist$datapath, sheet = ins[1], col_types = "text")
      colnames(indatasets)[1] = "Source"
      
      setProgress(0.2)
      
      for(i in seq(1:(length(ins)-1))) {
       
        indataset <- read_excel(Orderlist$datapath, sheet = ins[i+1],
                               col_types = "text")
        
        # Renaming the first column of the dataset to match what it is in the "Out" datasets, and to be able to bind the rows together for that column
        colnames(indataset)[1] <- "Source"
        
        indatasets <- bind_rows(indatasets, indataset)
        
        setProgress(0.2 + i/(length(ins) - 1)/100)
       
      }
      
      outshipping <- filter(outdatasets[, 1:18], !is.na(Source) | !is.na(`SO#`))
      outshipping <- filter(outshipping, !(Source %in% c("REPAIR", "Repair", "repair")))
      outshipping <- filter(outshipping, !startsWith(`SO#`, "RMA"))
      inorders <- filter(indatasets[, 1:11], !is.na(Source) | !(Source %in% c("REPAIR", "Repair", "repair")))
      inorders <- filter(inorders, !startsWith(SO, "RMA"))
      
      orders <- left_join(outshipping, inorders, by = c(`SO#` = "SO"))
      orders <- orders[!duplicated(orders), ]
      
      setProgress(0.3)
      
      orderslower <- data.frame(apply(orders, 2, tolower))
      
      commissionflag <- apply(orderslower, 1, function(row){
        any(grepl("commission", row))
      })
      refundflag <- apply(orderslower, 1, function(row){
        any(grepl("refund", row))
      })
      
      orders <- orders[!(refundflag | commissionflag), ]
      
      # This will take care of duplicates from matching many-to-many relationship of SO#s, while also preserving NET revenue.
      ordersgrouped <- orders %>%
        distinct(`SO#`, NET, .keep_all = TRUE)
      
      ordersgrouped <- mutate(ordersgrouped, productdetails = `Accessories/Details`)
      
      setProgress(0.4)
      
      ordersgrouped$Ordered <- as_date(as.numeric(ordersgrouped$Ordered) - 25569)
      
      ordersdates <- filter(ordersgrouped, !is.na(Ordered))
      # March 2029 orders
      ordersdates <- filter(ordersdates, !(`SO#` == "16147" | `SO#` == "16148"))
      ordersdates <- filter(ordersdates, Ordered > "2000-01-01")
      ordersdates$Ordered <- as_date(as.numeric(ordersdates$Ordered))
      
      productlist <- ordersdates$productdetails
      productlist2 <- ordersdates$Product
      
      productids <- str_match(productlist, "\\d{3}")
      productids2 <- str_match(productlist2, "\\d{3}")
      
      ordersdates <- ordersdates %>% 
        mutate(product_IDs = ifelse(is.na(productids), productids2, productids)) %>%
        select(product_IDs, everything())
      
      setProgress(0.5)
      
      ordersdates$product_IDs <- as.vector(ordersdates$product_IDs)
      
      cid <- c(202, 710, 602, 600, 203, 110, 340, 301)
      felix <- c(920, 950, 750, 751, 960, 940, 901, 900)
      interscan <- c(900)
      
      companycategories <- case_when(
        ordersdates$product_IDs %in% cid ~ "CID",
        tolower(ordersdates$Source.x) == "interscan" ~ "Interscan",
        tolower(ordersdates$Source.y) == "interscan" ~ "Interscan",
        ordersdates$product_IDs %in% felix ~ "Felix"
      )
      
      setProgress(0.6)
      
      ordersdates <- ordersdates %>%
        mutate(Company = companycategories) %>%
        mutate(YD = as.yearmon(Ordered), Year = year(Ordered)) %>%
        select(YD, everything()) %>%
        filter(YD != "Jan 2015") %>%
        mutate(wday = wday(Ordered, label = T, abbr = F)) %>%
        select(Company, wday, everything())

      SupportCases <- read_excel(Support$datapath)
      
      productdf <- tibble(unique(SupportCases$Product))
      productdf <- productdf %>%
       mutate(num = str_match(`unique(SupportCases$Product)`, "\\d{3}"))
      # This might be rough in terms of reproducibility for future updates to All Support Cases...
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
      
      setProgress(0.7)
      
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
      
      setProgress(0.8)
      
      ordersdates$product_IDs <- ifelse(!is.na(productids4), productids4, ordersdates$product_IDs)
      ordersdates$product_IDs <- ifelse(!is.na(productids3), productids3, ordersdates$product_IDs)
      
      ordersdates <- ordersdates %>%
        left_join(productdf, by = c("product_IDs" = "num")) %>%
        mutate(`END USER` = ifelse(`END USER` == "same", `CUSTOMER'S NAME`, `END USER`)) %>%
        mutate(customer = tolower(ifelse(is.na(`END USER`), `CUSTOMER'S NAME`, `END USER`))) %>%
        select(Product.y, everything()) %>%
        arrange(Ordered)
      
      pcustomers <- c()
      pcflag <- logical(nrow(ordersdates))  # Initialize pcflag as a logical vector
      
      for(i in 1:nrow(ordersdates)){
       
        if(ordersdates$customer[i] %in% pcustomers) {
        
          pcflag[i] <- TRUE
        
        } else {
        
          pcflag[i] <- FALSE
          pcustomers <- c(pcustomers, ordersdates$customer[i])
        
        }
       
      }
      
      ordersdates <- cbind(ordersdates, pcflag)
      
      output$output <- renderPrint({
        "Modified Orderlist ready for download"
      })
      
      setProgress(1)
      
      # Writes the csv to a file and downloads it
      output$downloadOrderlist <- downloadHandler(
        filename = "OrderlistModified.csv",
        content = function(file) {
          write_csv(ordersdates, file)
        }
      )
                   
    })
  })
  
  observeEvent(input$modifyBtn5, {
    
    PublishedPapers <- input$file6
    PapersTerms <- input$file7
    
    if(is.null(PublishedPapers) || is.null(PapersTerms)) {
      
      output$output <- renderPrint({
        "Please upload Published Papers and Published Papers Terms."
      })
      return()
      
    }
    
    withProgress(session = shiny::getDefaultReactiveDomain(),
                 message = "Processing files...",
                 detail = "Please wait...",
                 value = 0, {
      
      Papers110 <- as_tibble(read_excel(PublishedPapers$datapath, sheet = "Ci-110"))
      Papers202 <- as_tibble(read_excel(PublishedPapers$datapath, sheet = "CI-202"))
      Papers203 <- as_tibble(read_excel(PublishedPapers$datapath, sheet = "CI-203"))
      Papers340 <- as_tibble(read_excel(PublishedPapers$datapath, sheet = "CI-340"))
      Papers600 <- as_tibble(read_excel(PublishedPapers$datapath, sheet = "CI-600"))
      Papers602 <- as_tibble(read_excel(PublishedPapers$datapath, sheet = "CI-602"))
      Papers710 <- as_tibble(read_excel(PublishedPapers$datapath, sheet = "CI-710"))
      Papers203 <- rename(Papers203, "Abstract" = `...7`)
      Terms110 <- as_tibble(read_excel(PapersTerms$datapath, sheet = "CI-110"))
      Terms202203 <- as_tibble(read_excel(PapersTerms$datapath, sheet = "CI-202 203"))
      Terms340 <- as_tibble(read_excel(PapersTerms$datapath, sheet = "CI-340"))
      Terms600602 <- as_tibble(read_excel(PapersTerms$datapath, sheet = "CI-600 602"))
      Terms710 <- as_tibble(read_excel(PapersTerms$datapath, sheet = "CI-710S"))
      
      setProgress(0.2)
      
      Papers202$Date <- as.character(Papers202$Date)
      
      Papers202203 <- bind_rows(Papers202, Papers203)
      Papers600602 <- bind_rows(Papers600, Papers602)
      
      setProgress(0.4)
      
      PapersTransformation <- function(paper, terms) {
      
        paper <- paper %>%
          mutate(terms = tolower(gsub(" ", "", Abstract)))
        
        termstemp <- terms %>%
          rename("Terms1" = colnames(terms)[1]) %>%
          mutate(terms = tolower(gsub(" ", "", Terms1)))
        
        allabstracts <- paste(paper$terms, collapse = "")
        
        term_frequency <- c()
        
        for (i in 1:length(termstemp$terms)) {
        
          term_frequency[i] <- str_count(allabstracts, termstemp$terms[i])
        
        }
        
        result <- as_tibble(cbind(termstemp$Terms1, as.numeric(term_frequency), colnames(terms)[1]))
        
        result <- arrange(result, desc(term_frequency))
        
        return(result)
      
      }
      
      setProgress(0.6)
      
      Papers <- list(Papers110, Papers202203, Papers340, Papers600602, Papers710)
      Terms <- list(Terms110, Terms202203, Terms340, Terms600602, Terms710)
      
      papersterms <- tibble()
      
      for(i in 1:length(Papers)) {
      
        temppapers <- PapersTransformation(Papers[[i]], Terms[[i]])
        papersterms <- bind_rows(papersterms, temppapers)
      
      }
      
      setProgress(0.8)
      
      papersterms <- rename(papersterms, "Term" = "V1", "Frequency" = "V2", "Product" = "V3")
      papersterms <- na.omit(papersterms)
      
      setProgress(1)
      
      output$output <- renderPrint({
        "Published Papers ready for download"
      })
      
      output$downloadPapers <- downloadHandler(
        filename = "Papers Modified.csv",
        content = function(file) {
          write_csv(papersterms, file)
        }
      )
      
    })
  })
  
}



# Run the application 
shinyApp(ui = ui, server = server)

