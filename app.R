# Load required packages
library(shiny)
library(bs4Dash)
library(ggplot2)
library(DT)
library(readr)
library(readxl)
library(dplyr)
library(scales)
library(factoextra)
library(writexl)
library(plotly)
library(tidyr)

#------------------------------------------------------------------
# Helper Functions
#------------------------------------------------------------------

# Return expected risk-of-bias domain columns based on tool.
get_domain_cols <- function(data, tool) {
  if (tool == "ROB2") {
    return(intersect(names(data), c("Randomization", "Deviations", "Missing", "Measurement", "Selection")))
  } else if (tool == "ROBINS-I") {
    return(intersect(names(data), c("Confounding", "Selection", "Classification", "Deviations", "Missing", "Measurement", "Reporting")))
  } else if (tool == "QUADAS-2") {
    return(intersect(names(data), c("PatientSelection", "IndexTest", "ReferenceStandard", "FlowTiming")))
  } else if (tool == "ROB1") {
    return(intersect(names(data), c("RandomSequence", "AllocationConcealment", "BlindingParticipants", "BlindingOutcome", "IncompleteOutcome", "SelectiveReporting")))
  } else if (tool == "NOS") {
    return(intersect(names(data), c("Selection", "Comparability", "Outcome")))
  } else {
    # fallback if unknown tool
    return(setdiff(names(data), c("Study", "Overall", "Weight")))
  }
}

# Convert categorical judgments to numeric scores based on tool
convert_to_numeric <- function(x, tool) {
  if(tool == "ROB2") {
    ifelse(x == "Low", 1,
           ifelse(x == "Some concerns", 2,
                  ifelse(x == "High", 3, NA)))
  } else if(tool == "ROBINS-I") {
    ifelse(x == "Low", 1,
           ifelse(x == "Moderate", 2,
                  ifelse(x == "Serious", 3,
                         ifelse(x == "Critical", 4,
                                ifelse(x == "No information", 5, NA)))))
  } else if(tool == "QUADAS-2") {
    ifelse(x == "Low", 1,
           ifelse(x == "Unclear", 2,
                  ifelse(x == "High", 3, NA)))
  } else if(tool == "ROB1") {
    ifelse(x == "Low", 1,
           ifelse(x == "Unclear", 2,
                  ifelse(x == "High", 3, NA)))
  } else if(tool == "NOS") {
    ifelse(x == "Good", 1,
           ifelse(x == "Fair", 2,
                  ifelse(x == "Poor", 3, NA)))
  } else {
    NA
  }
}

# Summary Plot: Stacked bar chart
rob_summary <- function(data, tool, overall = FALSE) {
  domain_cols <- get_domain_cols(data, tool)
  plot_data <- data.frame()
  for (domain in domain_cols) {
    counts <- table(data[[domain]])
    percentages <- as.numeric(counts) / sum(counts)
    temp_df <- data.frame(
      Domain = domain,
      Level = names(counts),
      Percentage = percentages,
      stringsAsFactors = FALSE
    )
    plot_data <- rbind(plot_data, temp_df)
  }
  if (overall && "Overall" %in% names(data)) {
    counts <- table(data$Overall)
    percentages <- as.numeric(counts) / sum(counts)
    temp_df <- data.frame(
      Domain = "Overall",
      Level = names(counts),
      Percentage = percentages,
      stringsAsFactors = FALSE
    )
    plot_data <- rbind(plot_data, temp_df)
  }
  # Define palette based on tool
  if (tool == "ROB2") {
    palette <- c("Low" = "#66c2a5", "Some concerns" = "#fc8d62", "High" = "#8da0cb")
  } else if (tool == "ROBINS-I") {
    palette <- c("Low" = "#66c2a5", "Moderate" = "#fc8d62",
                 "Serious" = "#8da0cb", "Critical" = "#e78ac3",
                 "No information" = "#a6d854")
  } else if (tool == "QUADAS-2") {
    palette <- c("Low" = "#66c2a5", "High" = "#8da0cb", "Unclear" = "#fc8d62")
  } else if (tool == "ROB1") {
    palette <- c("Low" = "#66c2a5", "Unclear" = "#fc8d62", "High" = "#8da0cb")
  } else if (tool == "NOS") {
    palette <- c("Good" = "#66c2a5", "Fair" = "#fc8d62", "Poor" = "#8da0cb")
  } else {
    palette <- c("Low" = "#66c2a5", "Unclear" = "#fc8d62", "High" = "#8da0cb")
  }
  plot_data$Level <- factor(plot_data$Level, levels = names(palette))
  ggplot(plot_data, aes(x = Domain, y = Percentage, fill = Level)) +
    geom_bar(stat = "identity", position = "stack") +
    scale_fill_manual(values = palette, drop = FALSE) +
    theme_minimal() +
    labs(
      title = paste("Risk of Bias Summary -", tool),
      y = "Percentage", x = ""
    ) +
    scale_y_continuous(labels = percent_format(accuracy = 1))
}

# Traffic Light Plot: Dot plot per study
rob_traffic_light <- function(data, tool, psize = 10) {
  domain_cols <- get_domain_cols(data, tool)
  # Initialize with fixed columns so that rbind always gets the same structure
  plot_data <- data.frame(
    Study = character(),
    Domain = character(),
    Judgment = character(),
    stringsAsFactors = FALSE
  )
  for (i in 1:nrow(data)) {
    for (domain in domain_cols) {
      plot_data <- rbind(plot_data, data.frame(
        Study = data$Study[i],
        Domain = domain,
        Judgment = as.character(data[i, domain]),
        stringsAsFactors = FALSE
      ))
    }
    if ("Overall" %in% names(data)) {
      plot_data <- rbind(plot_data, data.frame(
        Study = data$Study[i],
        Domain = "Overall",
        Judgment = as.character(data$Overall[i]),
        stringsAsFactors = FALSE
      ))
    }
  }
  if (tool == "ROB2") {
    palette <- c("Low" = "#66c2a5", "Some concerns" = "#fc8d62", "High" = "#8da0cb")
  } else if (tool == "ROBINS-I") {
    palette <- c("Low" = "#66c2a5", "Moderate" = "#fc8d62",
                 "Serious" = "#8da0cb", "Critical" = "#e78ac3",
                 "No information" = "#a6d854")
  } else if (tool == "QUADAS-2") {
    palette <- c("Low" = "#66c2a5", "High" = "#8da0cb", "Unclear" = "#fc8d62")
  } else if (tool == "ROB1") {
    palette <- c("Low" = "#66c2a5", "Unclear" = "#fc8d62", "High" = "#8da0cb")
  } else if (tool == "NOS") {
    palette <- c("Good" = "#66c2a5", "Fair" = "#fc8d62", "Poor" = "#8da0cb")
  } else {
    palette <- c("Low" = "#66c2a5", "Unclear" = "#fc8d62", "High" = "#8da0cb")
  }
  plot_data$Judgment <- factor(plot_data$Judgment, levels = names(palette))
  ggplot(plot_data, aes(x = Domain, y = Study, color = Judgment)) +
    geom_point(size = psize/5) +
    scale_color_manual(values = palette, drop = FALSE) +
    theme_minimal() +
    labs(title = paste("Risk of Bias Traffic Light -", tool))
}

# Frequency Distribution Plot: Faceted bar charts per domain
rob_frequency_plot <- function(data, tool) {
  domain_cols <- get_domain_cols(data, tool)
  # Only pivot the expected risk-of-bias columns to avoid mixing types
  plot_data <- pivot_longer(
    data,
    cols = all_of(domain_cols),
    names_to = "Domain",
    values_to = "Judgment"
  )
  if (tool == "ROB2") {
    lev <- c("Low", "Some concerns", "High")
  } else if (tool == "ROBINS-I") {
    lev <- c("Low", "Moderate", "Serious", "Critical", "No information")
  } else if (tool == "QUADAS-2") {
    lev <- c("Low", "High", "Unclear")
  } else if (tool == "ROB1") {
    lev <- c("Low", "Unclear", "High")
  } else if (tool == "NOS") {
    lev <- c("Good", "Fair", "Poor")
  } else {
    lev <- NULL
  }
  if(!is.null(lev)) plot_data$Judgment <- factor(plot_data$Judgment, levels = lev)
  ggplot(plot_data, aes(x = Judgment, fill = Judgment)) +
    geom_bar() +
    facet_wrap(~Domain, scales = "free_y") +
    theme_minimal() +
    labs(
      title = paste("Frequency Distribution by Domain -", tool),
      x = "Judgment", y = "Count"
    )
}

# Cluster Analysis Plot
rob_cluster_analysis <- function(data, tool, numClusters = 3) {
  df <- data
  # Select only expected risk-of-bias columns
  domain_cols <- get_domain_cols(df, tool)
  # Convert categorical judgments to numeric scores for selected columns
  for (col in domain_cols) {
    df[[col]] <- convert_to_numeric(df[[col]], tool)
  }
  # Remove rows with missing data
  df_complete <- df[complete.cases(df[, domain_cols]), ]
  if(nrow(df_complete) < 2) return(NULL)
  km <- kmeans(df_complete[, domain_cols], centers = numClusters)
  df_complete$Cluster <- factor(km$cluster)
  ggplot(df_complete, aes_string(x = domain_cols[1], y = domain_cols[2],
                                 color = "Cluster", label = "Study")) +
    geom_point(size = 3) +
    geom_text(vjust = -0.5) +
    theme_minimal() +
    labs(title = paste("K-means Clustering (k =", numClusters, ")"))
}

#------------------------------------------------------------------
# Tables for Export
#------------------------------------------------------------------

get_summary_table <- function(data, tool, overall = FALSE) {
  domain_cols <- get_domain_cols(data, tool)
  summary_list <- lapply(domain_cols, function(domain) {
    cnts <- table(data[[domain]])
    pct <- round(100 * as.numeric(cnts) / sum(cnts), 1)
    data.frame(
      Domain = domain,
      Level = names(cnts),
      Count = as.numeric(cnts),
      Percentage = pct,
      stringsAsFactors = FALSE
    )
  })
  summary_df <- do.call(rbind, summary_list)
  if (overall && "Overall" %in% names(data)) {
    cnts <- table(data$Overall)
    pct <- round(100 * as.numeric(cnts) / sum(cnts), 1)
    overall_df <- data.frame(
      Domain = "Overall",
      Level = names(cnts),
      Count = as.numeric(cnts),
      Percentage = pct,
      stringsAsFactors = FALSE
    )
    summary_df <- rbind(summary_df, overall_df)
  }
  summary_df
}

get_traffic_table <- function(data, tool) {
  domain_cols <- get_domain_cols(data, tool)
  traffic_df <- data.frame(
    Study = character(),
    Domain = character(),
    Judgment = character(),
    stringsAsFactors = FALSE
  )
  for(i in 1:nrow(data)) {
    for(domain in domain_cols) {
      traffic_df <- rbind(
        traffic_df,
        data.frame(
          Study = data$Study[i],
          Domain = domain,
          Judgment = as.character(data[i, domain]),
          stringsAsFactors = FALSE
        )
      )
    }
    if("Overall" %in% names(data)) {
      traffic_df <- rbind(
        traffic_df,
        data.frame(
          Study = data$Study[i],
          Domain = "Overall",
          Judgment = as.character(data$Overall[i]),
          stringsAsFactors = FALSE
        )
      )
    }
  }
  traffic_df
}

get_frequency_table <- function(data, tool) {
  domain_cols <- get_domain_cols(data, tool)
  freq_list <- lapply(domain_cols, function(domain) {
    tab <- table(data[[domain]])
    data.frame(
      Domain = domain,
      Judgment = names(tab),
      Count = as.numeric(tab),
      stringsAsFactors = FALSE
    )
  })
  freq_df <- do.call(rbind, freq_list)
  freq_df
}

#------------------------------------------------------------------
# Sample Datasets
#------------------------------------------------------------------

create_rob2_data <- function() {
  set.seed(123)
  df <- data.frame(
    Study = paste0("Study_", 1:6),
    Randomization = sample(c("Low", "Some concerns", "High"), 6, replace = TRUE),
    Deviations = sample(c("Low", "Some concerns", "High"), 6, replace = TRUE),
    Missing = sample(c("Low", "Some concerns", "High"), 6, replace = TRUE),
    Measurement = sample(c("Low", "Some concerns", "High"), 6, replace = TRUE),
    Selection = sample(c("Low", "Some concerns", "High"), 6, replace = TRUE),
    stringsAsFactors = FALSE
  )
  df$Overall <- apply(df[, 2:6], 1, function(x) {
    if("High" %in% x) return("High")
    if("Some concerns" %in% x) return("Some concerns")
    return("Low")
  })
  df$Weight <- sample(50:200, 6, replace = TRUE)
  df
}

create_robins_data <- function() {
  set.seed(111)
  df <- data.frame(
    Study = paste0("Study_", 1:5),
    Confounding = sample(c("Low", "Moderate", "Serious", "Critical", "No information"), 5, replace = TRUE),
    Selection = sample(c("Low", "Moderate", "Serious", "Critical", "No information"), 5, replace = TRUE),
    Classification = sample(c("Low", "Moderate", "Serious", "Critical", "No information"), 5, replace = TRUE),
    Deviations = sample(c("Low", "Moderate", "Serious", "Critical", "No information"), 5, replace = TRUE),
    Missing = sample(c("Low", "Moderate", "Serious", "Critical", "No information"), 5, replace = TRUE),
    Measurement = sample(c("Low", "Moderate", "Serious", "Critical", "No information"), 5, replace = TRUE),
    Reporting = sample(c("Low", "Moderate", "Serious", "Critical", "No information"), 5, replace = TRUE),
    stringsAsFactors = FALSE
  )
  df$Overall <- apply(df[, 2:8], 1, function(x) {
    # The order determines the "worst" rating
    order <- c("Low", "Moderate", "Serious", "Critical", "No information")
    sorted <- match(x, order)
    order[max(sorted, na.rm = TRUE)]
  })
  df$Weight <- sample(100:300, 5, replace = TRUE)
  df
}

create_quadas_data <- function() {
  set.seed(222)
  df <- data.frame(
    Study = paste0("Study_", 1:6),
    PatientSelection = sample(c("Low", "High", "Unclear"), 6, replace = TRUE),
    IndexTest = sample(c("Low", "High", "Unclear"), 6, replace = TRUE),
    ReferenceStandard = sample(c("Low", "High", "Unclear"), 6, replace = TRUE),
    FlowTiming = sample(c("Low", "High", "Unclear"), 6, replace = TRUE),
    stringsAsFactors = FALSE
  )
  df$Overall <- apply(df[, 2:5], 1, function(x) {
    if("High" %in% x) return("High")
    if("Unclear" %in% x) return("Unclear")
    return("Low")
  })
  df$Weight <- sample(30:100, 6, replace = TRUE)
  df
}

create_rob1_data <- function() {
  set.seed(333)
  df <- data.frame(
    Study = paste0("Study_", 1:4),
    RandomSequence = sample(c("Low", "Unclear", "High"), 4, replace = TRUE),
    AllocationConcealment = sample(c("Low", "Unclear", "High"), 4, replace = TRUE),
    BlindingParticipants = sample(c("Low", "Unclear", "High"), 4, replace = TRUE),
    BlindingOutcome = sample(c("Low", "Unclear", "High"), 4, replace = TRUE),
    IncompleteOutcome = sample(c("Low", "Unclear", "High"), 4, replace = TRUE),
    SelectiveReporting = sample(c("Low", "Unclear", "High"), 4, replace = TRUE),
    stringsAsFactors = FALSE
  )
  df$Overall <- apply(df[, 2:7], 1, function(x) {
    if("High" %in% x) return("High")
    if("Unclear" %in% x) return("Unclear")
    return("Low")
  })
  df$Weight <- sample(20:50, 4, replace = TRUE)
  df
}

create_nos_data <- function() {
  set.seed(444)
  df <- data.frame(
    Study = paste0("Study_", 1:5),
    Selection = sample(c("Good", "Fair", "Poor"), 5, replace = TRUE),
    Comparability = sample(c("Good", "Fair", "Poor"), 5, replace = TRUE),
    Outcome = sample(c("Good", "Fair", "Poor"), 5, replace = TRUE),
    stringsAsFactors = FALSE
  )
  df$Overall <- apply(df[, 2:4], 1, function(x) {
    if("Poor" %in% x) return("Poor")
    if("Fair" %in% x) return("Fair")
    return("Good")
  })
  df$Weight <- sample(50:100, 5, replace = TRUE)
  df
}

# Create list of sample datasets for each tool
sample_datasets <- list(
  "ROB2" = create_rob2_data(),
  "ROBINS-I" = create_robins_data(),
  "QUADAS-2" = create_quadas_data(),
  "ROB1" = create_rob1_data(),
  "NOS" = create_nos_data()
)

#------------------------------------------------------------------
# UI
#------------------------------------------------------------------
ui <- bs4DashPage(
  title = "Multi Sample RoB App",
  header = bs4DashNavbar(
    title = tagList(icon("database"), "Multi-Sample RoB App")
  ),
  sidebar = bs4DashSidebar(
    skin = "light",
    status = "primary",
    bs4SidebarMenu(
      # 1) Data Management
      bs4SidebarMenuItem("Data Management", tabName = "data", icon = icon("upload")),
      # 2) Summary Plot
      bs4SidebarMenuItem("Summary Plot", tabName = "summary", icon = icon("chart-bar")),
      # 3) Traffic Light
      bs4SidebarMenuItem("Traffic Light", tabName = "traffic", icon = icon("traffic-light")),
      # 4) Frequency
      bs4SidebarMenuItem("Frequency", tabName = "frequency", icon = icon("table")),
      # 5) Analysis
      bs4SidebarMenuItem("Analysis", tabName = "analysis", icon = icon("chart-line")),
      # 6) Tables
      bs4SidebarMenuItem("Tables", tabName = "tables", icon = icon("file-alt")),
      # 7) Export & Report
      bs4SidebarMenuItem("Export & Report", tabName = "export", icon = icon("file-export"))
    )
  ),
  body = bs4DashBody(
    bs4TabItems(
      # Data Management Tab
      bs4TabItem(
        tabName = "data",
        fluidRow(
          column(
            width = 7,
            bs4Card(
              title = "Data Controls",
              status = "primary",
              solidHeader = TRUE,
              # Radio buttons for picking a sample dataset (including NOS)
              radioButtons("sampleChoice", "Pick a Sample Dataset:",
                           choices = c("ROB2", "ROBINS-I", "QUADAS-2", "ROB1", "NOS"),
                           selected = "ROB2"),
              actionButton("btnLoadSample", "Load Selected Sample", icon = icon("download")),
              hr(),
              downloadButton("dlSampleCSV", "Download Sample CSV", icon = icon("file-csv")),
              downloadButton("dlSampleExcel", "Download Sample Excel", icon = icon("file-excel")),
              hr(),
              radioButtons("fileType", "File Type",
                           choices = c("CSV" = "csv", "Excel" = "xlsx"),
                           selected = "csv"),
              fileInput("dataFile", "Upload a File", accept = c(".csv", ".xlsx")),
              hr(),
              
              # The selectInput for choosing the risk of bias tool (including NOS)
              selectInput("tool", "Risk of Bias Tool",
                          choices = c("ROB2", "ROBINS-I", "QUADAS-2", "ROB1","NOS" ),
                          selected = "ROB2"),
              helpText("After loading data, view it on the right.")
            )
          ),
          column(
            width = 6,
            bs4Card(
              title = "Data Table",
              status = "info",
              solidHeader = TRUE,
              DTOutput("dataTable")
            )
          )
        )
      ),
      # Summary Plot Tab
      bs4TabItem(
        tabName = "summary",
        fluidRow(
          bs4Card(
            title = "Summary Plot Options",
            status = "primary",
            solidHeader = TRUE,
            collapsible = TRUE,
            collapsed = TRUE,
            width = 12,
            checkboxInput("chkOverall", "Include Overall Column?", FALSE),
            selectInput("plotTheme", "Plot Theme",
                        choices = c("Minimal" = "theme_minimal", "Classic" = "theme_classic"),
                        selected = "theme_minimal"),
            downloadButton("dlSummaryPlot", "Download Summary Plot")
          )
        ),
        fluidRow(
          bs4Card(
            title = "Summary Plot",
            status = "primary",
            solidHeader = TRUE,
            width = 12,
            plotlyOutput("summaryPlot", height = "600px")
          )
        )
      ),
      # Traffic Light Tab
      bs4TabItem(
        tabName = "traffic",
        fluidRow(
          bs4Card(
            title = "Traffic Light Options",
            status = "success",
            solidHeader = TRUE,
            collapsible = TRUE,
            collapsed = TRUE,
            width = 12,
            sliderInput("sliderPtSize", "Point Size", min = 5, max = 30, value = 10),
            downloadButton("dlTrafficPlot", "Download Traffic Plot")
          )
        ),
        fluidRow(
          bs4Card(
            title = "Traffic Light Plot",
            status = "success",
            solidHeader = TRUE,
            width = 12,
            plotOutput("trafficPlot", height = "600px")
          )
        )
      ),
      # Frequency Distribution Tab
      bs4TabItem(
        tabName = "frequency",
        fluidRow(
          bs4Card(
            title = "Frequency Distribution Options",
            status = "primary",
            solidHeader = TRUE,
            collapsible = TRUE,
            collapsed = TRUE,
            width = 12,
            helpText("Displays the frequency of risk judgments for each domain."),
            downloadButton("dlFrequencyPlot", "Download Frequency Plot")
          )
        ),
        fluidRow(
          bs4Card(
            title = "Frequency Distribution Plot",
            status = "primary",
            solidHeader = TRUE,
            width = 12,
            plotOutput("frequencyPlot", height = "600px")
          )
        )
      ),
      # Analysis Tab
      bs4TabItem(
        tabName = "analysis",
        fluidRow(
          bs4Card(
            title = "Analysis Controls",
            status = "info",
            solidHeader = TRUE,
            collapsible = TRUE,
            collapsed = TRUE,
            width = 12,
            numericInput("numClusters", "Number of Clusters", 3, min = 2, max = 10),
            actionButton("btnCluster", "Run Clustering", icon = icon("play"))
          )
        ),
        fluidRow(
          bs4Card(
            title = "Cluster Analysis & Descriptive Stats",
            status = "info",
            solidHeader = TRUE,
            width = 12,
            plotOutput("clusterPlot", height = "500px"),
            hr(),
            verbatimTextOutput("statsOutput")
          )
        )
      ),
      # Tables Tab
      bs4TabItem(
        tabName = "tables",
        fluidRow(
          bs4Card(
            title = "Summary Table",
            status = "primary",
            solidHeader = TRUE,
            width = 12,
            DTOutput("summaryTable"),
            downloadButton("dlSummaryTable", "Download Summary Table", icon = icon("download"))
          )
        ),
        fluidRow(
          bs4Card(
            title = "Traffic Light Table",
            status = "success",
            solidHeader = TRUE,
            width = 12,
            DTOutput("trafficTable"),
            downloadButton("dlTrafficTable", "Download Traffic Table", icon = icon("download"))
          )
        ),
        fluidRow(
          bs4Card(
            title = "Frequency Table",
            status = "primary",
            solidHeader = TRUE,
            width = 12,
            DTOutput("frequencyTable"),
            downloadButton("dlFrequencyTable", "Download Frequency Table", icon = icon("download"))
          )
        )
      ),
      # Export & Report Tab
      bs4TabItem(
        tabName = "export",
        fluidRow(
          bs4Card(
            title = "Generate Text Report",
            status = "primary",
            solidHeader = TRUE,
            width = 12,
            downloadButton("btnReport", "Download Report")
          )
        )
      )
    )
  ),
  footer = bs4DashFooter(
    fixed = TRUE,
    left = "© 2025 Multi-Sample RoB App"
  )
)

#------------------------------------------------------------------
# Server
#------------------------------------------------------------------
server <- function(input, output, session) {
  
  # Reactive value for main data
  rvData <- reactiveVal(NULL)
  
  # 1) Load sample dataset when button is clicked
  observeEvent(input$btnLoadSample, {
    choice <- input$sampleChoice
    df <- sample_datasets[[choice]]
    rvData(df)
  })
  
  # 2) Download sample CSV (write on the fly)
  output$dlSampleCSV <- downloadHandler(
    filename = function() { paste0("sample_", input$sampleChoice, ".csv") },
    content = function(file) {
      write.csv(sample_datasets[[input$sampleChoice]], file, row.names = FALSE)
    }
  )
  
  # 3) Download sample Excel (write on the fly)
  output$dlSampleExcel <- downloadHandler(
    filename = function() { paste0("sample_", input$sampleChoice, ".xlsx") },
    content = function(file) {
      writexl::write_xlsx(sample_datasets[[input$sampleChoice]], file)
    }
  )
  
  # 4) Load user file upload
  observeEvent(input$dataFile, {
    req(input$dataFile)
    f <- input$dataFile$datapath
    if (input$fileType == "csv") {
      df <- read_csv(f, show_col_types = FALSE)
    } else {
      df <- readxl::read_excel(f)
    }
    rvData(df)
  })
  
  # 5) Create new template (default to ROB2 structure; adjust as needed)
  observeEvent(input$btnTemplate, {
    n <- 5
    domains <- c("Randomization", "Deviations", "Missing", "Measurement", "Selection")
    df <- data.frame(Study = paste0("Study_", 1:n))
    for (d in domains) df[[d]] <- NA_character_
    df$Overall <- NA_character_
    df$Weight  <- rep(1, n)
    rvData(df)
  })
  
  # Data Table
  output$dataTable <- renderDT({
    req(rvData())
    datatable(rvData(), options = list(pageLength = 10, scrollX = TRUE))
  })
  
  # Summary Plot (Interactive via Plotly)
  output$summaryPlot <- renderPlotly({
    req(rvData())
    p <- rob_summary(rvData(), tool = input$tool, overall = input$chkOverall)
    if (!is.null(input$plotTheme)) {
      theme_func <- get(input$plotTheme, asNamespace("ggplot2"))
      p <- p + theme_func()
    }
    ggplotly(p)
  })
  
  output$dlSummaryPlot <- downloadHandler(
    filename = function() { "summary_plot.png" },
    content = function(file) {
      ggsave(
        file,
        rob_summary(rvData(), tool = input$tool, overall = input$chkOverall),
        width = 8, height = 4, dpi = 300
      )
    }
  )
  
  # Traffic Light Plot
  output$trafficPlot <- renderPlot({
    req(rvData())
    rob_traffic_light(rvData(), tool = input$tool, psize = input$sliderPtSize)
  })
  
  output$dlTrafficPlot <- downloadHandler(
    filename = function() { "traffic_light_plot.png" },
    content = function(file) {
      ggsave(
        file,
        rob_traffic_light(rvData(), tool = input$tool, psize = input$sliderPtSize),
        width = 8, height = 6, dpi = 300
      )
    }
  )
  
  # Frequency Distribution Plot
  output$frequencyPlot <- renderPlot({
    req(rvData())
    rob_frequency_plot(rvData(), tool = input$tool)
  })
  
  output$dlFrequencyPlot <- downloadHandler(
    filename = function() { "frequency_plot.png" },
    content = function(file) {
      ggsave(
        file,
        rob_frequency_plot(rvData(), tool = input$tool),
        width = 10, height = 6, dpi = 300
      )
    }
  )
  
  # Analysis: Cluster Analysis and Descriptive Stats
  clusterResult <- eventReactive(input$btnCluster, {
    req(rvData())
    rob_cluster_analysis(rvData(), tool = input$tool, numClusters = input$numClusters)
  })
  
  output$clusterPlot <- renderPlot({
    clusterResult()
  })
  
  output$statsOutput <- renderPrint({
    req(rvData())
    df <- rvData()
    cat("Number of studies:", nrow(df), "\n")
    if ("Overall" %in% names(df)) {
      cat("Overall distribution:\n")
      print(table(df$Overall))
    } else {
      cat("No Overall column found.\n")
    }
  })
  
  # Tables Tab
  output$summaryTable <- renderDT({
    req(rvData())
    datatable(
      get_summary_table(rvData(), tool = input$tool, overall = input$chkOverall),
      options = list(pageLength = 10, scrollX = TRUE)
    )
  })
  
  output$trafficTable <- renderDT({
    req(rvData())
    datatable(
      get_traffic_table(rvData(), tool = input$tool),
      options = list(pageLength = 10, scrollX = TRUE)
    )
  })
  
  output$frequencyTable <- renderDT({
    req(rvData())
    datatable(
      get_frequency_table(rvData(), tool = input$tool),
      options = list(pageLength = 10, scrollX = TRUE)
    )
  })
  
  output$dlSummaryTable <- downloadHandler(
    filename = function() { "summary_table.csv" },
    content = function(file) {
      write.csv(
        get_summary_table(rvData(), tool = input$tool, overall = input$chkOverall),
        file, row.names = FALSE
      )
    }
  )
  
  output$dlTrafficTable <- downloadHandler(
    filename = function() { "traffic_table.csv" },
    content = function(file) {
      write.csv(get_traffic_table(rvData(), tool = input$tool), file, row.names = FALSE)
    }
  )
  
  output$dlFrequencyTable <- downloadHandler(
    filename = function() { "frequency_table.csv" },
    content = function(file) {
      write.csv(get_frequency_table(rvData(), tool = input$tool), file, row.names = FALSE)
    }
  )
  
  # Export Report (Text)
  output$btnReport <- downloadHandler(
    filename = function() { "rob_report.txt" },
    content = function(file) {
      df <- rvData()
      if (is.null(df)) {
        cat("No data loaded.\n", file = file)
        return()
      }
      cat("Risk of Bias Report\n", file = file)
      cat("-------------------\n", file = file, append = TRUE)
      cat("Number of studies:", nrow(df), "\n", file = file, append = TRUE)
      if ("Overall" %in% names(df)) {
        cat("Overall distribution:\n", file = file, append = TRUE)
        cat(capture.output(print(table(df$Overall))), sep = "\n", file = file, append = TRUE)
      }
    }
  )
}

shinyApp(ui, server)
