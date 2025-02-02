################################################################################
#####################  Packages  ###############################################
################################################################################

# Install pacman for package management if not already installed
if (!requireNamespace("pacman", quietly = TRUE)) {
  install.packages("pacman")
}
library(pacman)

# Required CRAN packages
cran_packages <- c(
  "readxl", "writexl", "tidyverse", "ggplot2", "ComplexUpset", 
  "ggVennDiagram", "tibble", "UpSetR", "ComplexUpset",
  "ComplexHeatmap" # Added ggVennDiagram
  
)

# Required Bioconductor packages
bioc_packages <- c(
  "VennDetail"
)

# Automatically install and load CRAN and Bioconductor packages
p_load(char = cran_packages, install = TRUE)
p_load(char = bioc_packages, install = TRUE, repos = BiocManager::repositories())

################################################################################
#####################  Set Current Working Directory ###########################
################################################################################

current_path <- rstudioapi::getActiveDocumentContext()$path 
setwd(dirname(current_path))
print(current_path)
base_dir <- dirname(current_path)

# Create output directory, if it doesn't exist
output_dir <- file.path(base_dir, "VennOutput")
dir.create(output_dir, showWarnings = FALSE, recursive = TRUE)



################################################################################
#####################  Custom functions        #################################
################################################################################

# detach("package:VennDetail", unload = TRUE)
# detach("package:UpSetR", unload = TRUE)
# library(UpSetR)


# This function automatically find the appropriate color index
# based on the remainder,as the index, of i divided by the length of mycol
get_color_by_row <- function (mycol, i) {
  remainder <- i %% length(mycol)
  if (remainder == 0) {
    remainder <- length(mycol)
  }
  return(as.character(mycol[remainder])) 
}


## The following functions are from VennDetail and UpSetR packages
## We modified some of these function to adjust the colors of 
## the intersection part

## updated Create_layout from UpSetR 
Create_layout <- function(setup, mat_color, mat_col, matrix_dot_alpha,
                          mycol=NULL, intersect.mycol.palette.use=FALSE){
  Matrix_layout <- expand.grid(y=seq(nrow(setup)), x=seq(ncol(setup)))
  Matrix_layout <- data.frame(Matrix_layout, value = as.vector(setup))
  for(i in 1:nrow(Matrix_layout)){
    if(Matrix_layout$value[i] > as.integer(0)){
      if (intersect.mycol.palette.use) {
        if (is.null(mycol)) {
          Matrix_layout$color[i] <- mat_color
        } else {
          Matrix_layout$color[i] <- get_color_by_row(mycol, i)
        }
      } else {
        Matrix_layout$color[i] <- mat_color
      }
      
      Matrix_layout$alpha[i] <- 1
      Matrix_layout$Intersection[i] <- paste(Matrix_layout$x[i], "yes", sep ="")
    }
    else{
      
      Matrix_layout$color[i] <- "gray83"
      Matrix_layout$alpha[i] <- matrix_dot_alpha
      Matrix_layout$Intersection[i] <- paste(i, "No", sep = "")
    }
  }
  if(is.null(mat_col) == F){
    for(i in 1:nrow(mat_col)){
      mat_x <- mat_col$x[i]
      mat_color <- as.character(mat_col$color[i])
      for(i in 1:nrow(Matrix_layout)){
        if((Matrix_layout$x[i] == mat_x) && (Matrix_layout$value[i] != 0)){
          Matrix_layout$color[i] <- mat_color
          #message(mat_color)
        }
      }
    }
  }
  return(Matrix_layout)
}


## updated plot.Venn from VennDetail
plot.Venn <- function(x, type = "venn", col = "black", sep = "_",
                      mycol = NULL, cat.cex = 1.5, alpha = 0.5, cex = 2,
                      cat.fontface = "bold",
                      margin = 0.05, text.scale = c(1.5, 1.5, 1.5, 1.5, 1.5, 1.5),
                      filename = NULL, piecolor = NULL, revcolor = "lightgrey",
                      any = NULL, show.number = TRUE, show.x = TRUE, log = FALSE,
                      base = NULL, percentage = FALSE, sets.x.label = "Set Size",
                      mainbar.y.label = "Intersection Size", nintersects = 40,
                      abbr= FALSE, abbr.method = "both.sides", minlength = 3, 
                      intersect.line.color = "gray23", 
                      intersect.mycol.palette.use = FALSE, ...)
{
  result <- x
  x <- x@input
  
  if (is.null(mycol)) {
    mycol = c("dodgerblue", "goldenrod1", "darkorange1", "seagreen3", "orchid3")
  } else if (length(mycol)<5) {
    temp_colors <- c("dodgerblue", "goldenrod1", "darkorange1", "seagreen3", "orchid3")
    temp_colors[1:length(mycol)] <- mycol 
    mycol <- temp_colors
  }

  if(type == "venn"&&length(x) <= 5){
    #require(VennDiagram)
    n <- length(x)
    ## don't generate log, except on warning / error.
    othresh <- futile.logger::flog.threshold()
    futile.logger::flog.threshold(futile.logger::WARN)
    on.exit(futile.logger::flog.threshold(othresh))
    p <- venn.diagram(x, filename = filename,
                      col = col,
                      fill = mycol[seq_len(n)],
                      alpha = alpha,
                      cex = cex,
                      cat.col = mycol[seq_len(n)],
                      cat.cex = cat.cex,
                      cat.fontface = cat.fontface,
                      #cat.pos=cat.pos,
                      #cat.dist=cat.dist,
                      margin = margin)
    grid.draw(p)
  }
  if(length(x) > 5 & type != "upset"){
    type <- "vennpie"
  }
  if(type == "vennpie"){
    print(vennpie(result, sep = sep, color = piecolor, revcolor = revcolor,
                  any = any, show.number = show.number, show.x = show.x, log = log,
                  base = base, percentage = percentage))
  }
  if(type == "upset"){
    if ((is.null(mycol)) || (length(x) > length(mycol))){
      upset(fromList(x), nsets = length(x), sets.x.label = sets.x.label,
            mainbar.y.label = mainbar.y.label, nintersects = nintersects,
            point.size = 5, sets.bar.color = setcolor(length(x)),
            text.scale = text.scale)
    }else{
      upset(fromList(x), nsets = length(x), sets.x.label = sets.x.label,
            mainbar.y.label = mainbar.y.label, nintersects = nintersects,
            point.size = 5, sets.bar.color = mycol[seq_along(x)],
            text.scale = text.scale, intersect.line.color=intersect.line.color,
            intersect.mycol.palette.use=intersect.mycol.palette.use,
            mycol=mycol)
    }
  }
}


## Updated upset function from UpSetR
upset <- function(data, nsets = 5, nintersects = 40, sets = NULL, keep.order = F, set.metadata = NULL, intersections = NULL,
                  matrix.color = "gray23", main.bar.color = "gray23", mainbar.y.label = "Intersection Size", mainbar.y.max = NULL,
                  sets.bar.color = "gray23", plot.title = NA, sets.x.label = "Set Size", point.size = 2.2, line.size = 0.7,
                  mb.ratio = c(0.70,0.30), expression = NULL, att.pos = NULL, att.color = main.bar.color, order.by = c("freq", "degree"),
                  decreasing = c(T, F), show.numbers = "yes", number.angles = 0, number.colors=NULL, group.by = "degree",cutoff = NULL,
                  queries = NULL, query.legend = "none", shade.color = "gray88", shade.alpha = 0.25, matrix.dot.alpha =0.5,
                  empty.intersections = NULL, color.pal = 1, boxplot.summary = NULL, attribute.plots = NULL, scale.intersections = "identity",
                  scale.sets = "identity", text.scale = 1, set_size.angles = 0 , set_size.show = FALSE, set_size.numbers_size = NULL, set_size.scale_max = NULL,
                  intersect.line.color = "gray83", mycol = NULL,
                  intersect.mycol.palette.use = FALSE){
  
  # for custom intersection plots
  if (is.null(mycol)) {
    mycol = c("dodgerblue", "goldenrod1", "darkorange1", "seagreen3", "orchid3")
  } else if (length(mycol)<5) {
    temp_colors <- c("dodgerblue", "goldenrod1", "darkorange1", "seagreen3", "orchid3")
    temp_colors[1:length(mycol)] <- mycol 
    mycol <- temp_colors
  }
  
  startend <-UpSetR:::FindStartEnd(data)
  first.col <- startend[1]
  last.col <- startend[2]
  
  if(color.pal == 1){
    palette <- c("#1F77B4", "#FF7F0E", "#2CA02C", "#D62728", "#9467BD", "#8C564B", "#E377C2",
                 "#7F7F7F", "#BCBD22", "#17BECF")
  }
  else{
    palette <- c("#E69F00", "#56B4E9", "#009E73", "#F0E442", "#0072B2", "#D55E00",
                 "#CC79A7")
  }
  
  if(is.null(intersections) == F){
    Set_names <- unique((unlist(intersections)))
    Sets_to_remove <- UpSetR:::Remove(data, first.col, last.col, Set_names)
    New_data <- UpSetR:::Wanted(data, Sets_to_remove)
    Num_of_set <- UpSetR:::Number_of_sets(Set_names)
    if(keep.order == F){
      Set_names <- UpSetR:::order_sets(New_data, Set_names)
    }
    All_Freqs <- UpSetR:::specific_intersections(data, first.col, last.col, intersections, order.by, group.by, decreasing,
                                                 cutoff, main.bar.color, Set_names)
  }
  else if(is.null(intersections) == T){
    Set_names <- sets
    if(is.null(Set_names) == T || length(Set_names) == 0 ){
      Set_names <- UpSetR:::FindMostFreq(data, first.col, last.col, nsets)
    }
    Sets_to_remove <- UpSetR:::Remove(data, first.col, last.col, Set_names)
    New_data <- UpSetR:::Wanted(data, Sets_to_remove)
    Num_of_set <- UpSetR:::Number_of_sets(Set_names)
    if(keep.order == F){
      Set_names <- UpSetR:::order_sets(New_data, Set_names)
    }
    All_Freqs <- UpSetR:::Counter(New_data, Num_of_set, first.col, Set_names, nintersects, main.bar.color,
                                  order.by, group.by, cutoff, empty.intersections, decreasing)
  }
  Matrix_setup <- UpSetR:::Create_matrix(All_Freqs)
  labels <- UpSetR:::Make_labels(Matrix_setup)
  #Chose NA to represent NULL case as result of NA being inserted when at least one contained both x and y
  #i.e. if one custom plot had both x and y, and others had only x, the y's for the other plots were NA
  #if I decided to make the NULL case (all x and no y, or vice versa), there would have been alot more if/else statements
  #NA can be indexed so that we still get the non NA y aesthetics on correct plot. NULL cant be indexed.
  att.x <- c(); att.y <- c();
  if(is.null(attribute.plots) == F){
    for(i in seq_along(attribute.plots$plots)){
      if(length(attribute.plots$plots[[i]]$x) != 0){
        att.x[i] <- attribute.plots$plots[[i]]$x
      }
      else if(length(attribute.plots$plots[[i]]$x) == 0){
        att.x[i] <- NA
      }
      if(length(attribute.plots$plots[[i]]$y) != 0){
        att.y[i] <- attribute.plots$plots[[i]]$y
      }
      else if(length(attribute.plots$plots[[i]]$y) == 0){
        att.y[i] <- NA
      }
    }
  }
  
  BoxPlots <- NULL
  if(is.null(boxplot.summary) == F){
    BoxData <- UpSetR:::IntersectionBoxPlot(All_Freqs, New_data, first.col, Set_names)
    BoxPlots <- list()
    for(i in seq_along(boxplot.summary)){
      BoxPlots[[i]] <- UpSetR:::BoxPlotsPlot(BoxData, boxplot.summary[i], att.color)
    }
  }
  
  customAttDat <- NULL
  customQBar <- NULL
  Intersection <- NULL
  Element <- NULL
  legend <- NULL
  EBar_data <- NULL
  if(is.null(queries) == F){
    custom.queries <- UpSetR:::SeperateQueries(queries, 2, palette)
    customDat <- UpSetR:::customQueries(New_data, custom.queries, Set_names)
    legend <- UpSetR:::GuideGenerator(queries, palette)
    legend <- UpSetR:::Make_legend(legend)
    if(is.null(att.x) == F && is.null(customDat) == F){
      customAttDat <- UpSetR:::CustomAttData(customDat, Set_names)
    }
    customQBar <- UpSetR:::customQueriesBar(customDat, Set_names, All_Freqs, custom.queries)
  }
  if(is.null(queries) == F){
    Intersection <- UpSetR:::SeperateQueries(queries, 1, palette)
    Matrix_col <- intersects(UpSetR:::QuerieInterData, Intersection, New_data, first.col, Num_of_set,
                             All_Freqs, expression, Set_names, palette)
    Element <- UpSetR:::SeperateQueries(queries, 1, palette)
    EBar_data <-UpSetR:::ElemBarDat(Element, New_data, first.col, expression, Set_names,palette, All_Freqs)
  }
  else{
    Matrix_col <- NULL
  }
  
  Matrix_layout <- UpSetR:::Create_layout(Matrix_setup, matrix.color, Matrix_col, 
                                          matrix.dot.alpha, mycol, intersect.mycol.palette.use)
  Set_sizes <- UpSetR:::FindSetFreqs(New_data, first.col, Num_of_set, Set_names, keep.order)
  Bar_Q <- NULL
  if(is.null(queries) == F){
    Bar_Q <- intersects(UpSetR:::QuerieInterBar, Intersection, New_data, first.col, Num_of_set, All_Freqs, expression, Set_names, palette)
  }
  QInter_att_data <- NULL
  QElem_att_data <- NULL
  if((is.null(queries) == F) & (is.null(att.x) == F)){
    QInter_att_data <- intersects(UpSetR:::QuerieInterAtt, Intersection, New_data, first.col, Num_of_set, att.x, att.y,
                                  expression, Set_names, palette)
    QElem_att_data <- elements(UpSetR:::QuerieElemAtt, Element, New_data, first.col, expression, Set_names, att.x, att.y,
                               palette)
  }
  AllQueryData <- UpSetR:::combineQueriesData(QInter_att_data, QElem_att_data, customAttDat, att.x, att.y)
  
  ShadingData <- NULL
  
  if(is.null(set.metadata) == F){
    ShadingData <- UpSetR:::get_shade_groups(set.metadata, Set_names, Matrix_layout, shade.alpha)
    output <- UpSetR:::Make_set_metadata_plot(set.metadata, Set_names)
    set.metadata.plots <- output[[1]]
    set.metadata <- output[[2]]
    
    if(is.null(ShadingData) == FALSE){
      shade.alpha <- unique(ShadingData$alpha)
    }
  } else {
    set.metadata.plots <- NULL
  }
  if(is.null(ShadingData) == TRUE){
    ShadingData <- UpSetR:::MakeShading(Matrix_layout, shade.color)
  }
  Main_bar <- suppressMessages(UpSetR:::Make_main_bar(All_Freqs, Bar_Q, show.numbers, mb.ratio, customQBar, number.angles, number.colors, EBar_data, mainbar.y.label,
                                                      mainbar.y.max, scale.intersections, text.scale, attribute.plots, plot.title))
  Matrix <- UpSetR:::Make_matrix_plot(Matrix_layout, Set_sizes, All_Freqs, point.size, line.size,
                                      text.scale, labels, ShadingData, shade.alpha,intersect.line.color=intersect.line.color)
  Sizes <- UpSetR:::Make_size_plot(Set_sizes, sets.bar.color, mb.ratio, sets.x.label, scale.sets, text.scale, set_size.angles,set_size.show,
                                   set_size.scale_max, set_size.numbers_size)
  
  # Make_base_plot(Main_bar, Matrix, Sizes, labels, mb.ratio, att.x, att.y, New_data,
  #                expression, att.pos, first.col, att.color, AllQueryData, attribute.plots,
  #                legend, query.legend, BoxPlots, Set_names, set.metadata, set.metadata.plots)
  
  structure(class = "upset",
            .Data=list(
              Main_bar = Main_bar,
              Matrix = Matrix,
              Sizes = Sizes,
              labels = labels,
              mb.ratio = mb.ratio,
              att.x = att.x,
              att.y = att.y,
              New_data = New_data,
              expression = expression,
              att.pos = att.pos,
              first.col = first.col,
              att.color = att.color,
              AllQueryData = AllQueryData,
              attribute.plots = attribute.plots,
              legend = legend,
              query.legend = query.legend,
              BoxPlots = BoxPlots,
              Set_names = Set_names,
              set.metadata = set.metadata,
              set.metadata.plots = set.metadata.plots)
  )
}


## Updated Make_main_bar function from UpSetR
Make_main_bar <- function(Main_bar_data, Q, show_num, ratios, customQ, number_angles, number.colors,
                          ebar, ylabel, ymax, scale_intersections, text_scale, attribute_plots, plot.title){
  
  bottom_margin <- (-1)*0.65
  
  if(is.null(attribute_plots) == FALSE){
    bottom_margin <- (-1)*0.45
  }
  
  if(length(text_scale) > 1 && length(text_scale) <= 6){
    y_axis_title_scale <- text_scale[1]
    y_axis_tick_label_scale <- text_scale[2]
    intersection_size_number_scale <- text_scale[6]
  }
  else{
    y_axis_title_scale <- text_scale
    y_axis_tick_label_scale <- text_scale
    intersection_size_number_scale <- text_scale
  }
  
  if(is.null(Q) == F){
    inter_data <- Q
    if(nrow(inter_data) != 0){
      inter_data <- inter_data[order(inter_data$x), ]
    }
    else{inter_data <- NULL}
  }
  else{inter_data <- NULL}
  
  if(is.null(ebar) == F){
    elem_data <- ebar
    if(nrow(elem_data) != 0){
      elem_data <- elem_data[order(elem_data$x), ]
    }
    else{elem_data <- NULL}
  }
  else{elem_data <- NULL}
  
  #ten_perc creates appropriate space above highest bar so number doesnt get cut off
  if(is.null(ymax) == T){
    ten_perc <- ((max(Main_bar_data$freq)) * 0.1)
    ymax <- max(Main_bar_data$freq) + ten_perc
  }
  
  if(ylabel == "Intersection Size" && scale_intersections != "identity"){
    ylabel <- paste("Intersection Size", paste0("( ", scale_intersections, " )"))
  }
  if(scale_intersections == "log2"){
    Main_bar_data$freq <- round(log2(Main_bar_data$freq), 2)
    ymax <- log2(ymax)
  }
  if(scale_intersections == "log10"){
    Main_bar_data$freq <- round(log10(Main_bar_data$freq), 2)
    ymax <- log10(ymax)
  }
  Main_bar_plot <- (ggplot(data = Main_bar_data, aes_string(x = "x", y = "freq")) 
                    + scale_y_continuous(trans = scale_intersections)
                    + ylim(0, ymax)
                    + geom_bar(stat = "identity", width = 0.6,
                               fill = Main_bar_data$color)
                    + scale_x_continuous(limits = c(0,(nrow(Main_bar_data)+1 )), expand = c(0,0),
                                         breaks = NULL)
                    + xlab(NULL) + ylab(ylabel) +labs(title = NULL)
                    + theme(panel.background = element_rect(fill = "white"),
                            plot.margin = unit(c(0.5,0.5,bottom_margin,0.5), "lines"), panel.border = element_blank(),
                            axis.title.y = element_text(vjust = -0.8, size = 8.3*y_axis_title_scale), axis.text.y = element_text(vjust=0.3,
                                                                                                                                 size=7*y_axis_tick_label_scale)))
  if((show_num == "yes") || (show_num == "Yes")){
    if(is.null(number.colors)) {
      Main_bar_plot <- (Main_bar_plot + geom_text(aes_string(label = "freq"), size = 2.2*intersection_size_number_scale, vjust = -1,
                                                  angle = number_angles, colour = Main_bar_data$color))
    } else {
      Main_bar_plot <- (Main_bar_plot + geom_text(aes_string(label = "freq"), size = 2.2*intersection_size_number_scale, vjust = -1,
                                                  angle = number_angles, colour = number.colors))
    }
  }
  
  bInterDat <- NULL
  pInterDat <- NULL
  bCustomDat <- NULL
  pCustomDat <- NULL
  bElemDat <- NULL
  pElemDat <- NULL
  if(is.null(elem_data) == F){
    bElemDat <- elem_data[which(elem_data$act == T), ]
    bElemDat <- bElemDat[order(bElemDat$x), ]
    pElemDat <- elem_data[which(elem_data$act == F), ]
  }
  if(is.null(inter_data) == F){
    bInterDat <- inter_data[which(inter_data$act == T), ]
    bInterDat <- bInterDat[order(bInterDat$x), ]
    pInterDat <- inter_data[which(inter_data$act == F), ]
  }
  if(length(customQ) != 0){
    pCustomDat <- customQ[which(customQ$act == F), ]
    bCustomDat <- customQ[which(customQ$act == T), ]
    bCustomDat <- bCustomDat[order(bCustomDat$x), ]
  }
  if(length(bInterDat) != 0){
    Main_bar_plot <- Main_bar_plot + geom_bar(data = bInterDat,
                                              aes_string(x="x", y = "freq"),
                                              fill = bInterDat$color,
                                              stat = "identity", position = "identity", width = 0.6)
  }
  if(length(bElemDat) != 0){
    Main_bar_plot <- Main_bar_plot + geom_bar(data = bElemDat,
                                              aes_string(x="x", y = "freq"),
                                              fill = bElemDat$color,
                                              stat = "identity", position = "identity", width = 0.6)
  }
  if(length(bCustomDat) != 0){
    Main_bar_plot <- (Main_bar_plot + geom_bar(data = bCustomDat, aes_string(x="x", y = "freq2"),
                                               fill = bCustomDat$color2,
                                               stat = "identity", position ="identity", width = 0.6))
  }
  if(length(pCustomDat) != 0){
    Main_bar_plot <- (Main_bar_plot + geom_point(data = pCustomDat, aes_string(x="x", y = "freq2"), colour = pCustomDat$color2,
                                                 size = 2, shape = 17, position = position_jitter(width = 0.2, height = 0.2)))
  }
  if(length(pInterDat) != 0){
    Main_bar_plot <- (Main_bar_plot + geom_point(data = pInterDat, aes_string(x="x", y = "freq"),
                                                 position = position_jitter(width = 0.2, height = 0.2),
                                                 colour = pInterDat$color, size = 2, shape = 17))
  }
  if(length(pElemDat) != 0){
    Main_bar_plot <- (Main_bar_plot + geom_point(data = pElemDat, aes_string(x="x", y = "freq"),
                                                 position = position_jitter(width = 0.2, height = 0.2),
                                                 colour = pElemDat$color, size = 2, shape = 17))
  }
  
  Main_bar_plot <- (Main_bar_plot 
                    + geom_vline(xintercept = 0, color = "gray0")
                    + geom_hline(yintercept = 0, color = "gray0"))
  
  if(!is.na(plot.title)) {
    Main_bar_plot <- c(Main_bar_plot + ggtitle(plot.title))
  }
  Main_bar_plot <- ggplotGrob(Main_bar_plot)
  return(Main_bar_plot)
}


## Updated Make_matrix_plot function from UpSetR
Make_matrix_plot <- function(Mat_data,Set_size_data, Main_bar_data, point_size, line_size, text_scale, labels,
                             shading_data, shade_alpha, intersect.line.color="gray83"){
  
  if(length(text_scale) == 1){
    name_size_scale <- text_scale
  }
  if(length(text_scale) > 1 && length(text_scale) <= 6){
    name_size_scale <- text_scale[5]
  }
  
  #message(Mat_data$color)
  
  Matrix_plot <- (ggplot()
                  + theme(panel.background = element_rect(fill = "white"),
                          plot.margin=unit(c(-0.2,0.5,0.5,0.5), "lines"),
                          axis.text.x = element_blank(),
                          axis.ticks.x = element_blank(),
                          axis.ticks.y = element_blank(),
                          axis.text.y = element_text(colour = "gray0",
                                                     size = 7*name_size_scale, hjust = 0.4),
                          panel.grid.major = element_blank(),
                          panel.grid.minor = element_blank())
                  + xlab(NULL) + ylab("   ")
                  + scale_y_continuous(breaks = c(1:nrow(Set_size_data)),
                                       limits = c(0.5,(nrow(Set_size_data) +0.5)),
                                       labels = labels, expand = c(0,0))
                  + scale_x_continuous(limits = c(0,(nrow(Main_bar_data)+1 )), expand = c(0,0))
                  + geom_rect(data = shading_data, aes_string(xmin = "min", xmax = "max",
                                                              ymin = "y_min", ymax = "y_max"),
                              fill = shading_data$shade_color, alpha = shade_alpha)
                  + geom_point(data= Mat_data, aes_string(x= "x", y= "y"), colour = Mat_data$color,
                               size= point_size, alpha = Mat_data$alpha, shape=16)
                  + geom_line(data= Mat_data, aes_string(group = "Intersection", x="x", y="y"),
                                                         colour = intersect.line.color, size = line_size)
                  + scale_color_identity())
  Matrix_plot <- ggplot_gtable(ggplot_build(Matrix_plot))
  return(Matrix_plot)
}


## Assign the updated functions to relevant packages in the current workspace
assignInNamespace(x="Make_matrix_plot", value=Make_matrix_plot, ns="UpSetR")
assignInNamespace(x="Create_layout", value=Create_layout, ns="UpSetR")
assignInNamespace(x="plot.Venn", value=plot.Venn, ns="VennDetail")
assignInNamespace(x="upset", value=upset, ns="UpSetR")
assignInNamespace(x="Make_main_bar", value=Make_main_bar, ns="UpSetR")


# plot(six_way_venn, type="upset", mycol=custom_colors, intersect.line.color="gray30")
# plot(six_way_venn, type="upset", mycol=custom_colors, intersect.line.color="red")
# plot(six_way_venn, type="upset", mycol=custom_colors, 
#      intersect.mycol.palette.use=TRUE)





################################################################################
#####################  Define local venn diagram process function 
################################################################################

# Function to process data and create visualizations
process_venn_diagrams <- function(file_path, prefix) {
  
  # Read all six worksheets
  pfizer_data <- read_excel(file_path, sheet = "PFIZER-BIONTECH")
  pfizer_bivalent_data <- read_excel(file_path, sheet = "PFIZER-BIONTECH BIVALENT")
  moderna_data <- read_excel(file_path, sheet = "MODERNA")
  moderna_bivalent_data <- read_excel(file_path, sheet = "MODERNA BIVALENT")
  janssen_data <- read_excel(file_path, sheet = "JANSSEN")
  novavax_data <- read_excel(file_path, sheet = "NOVAVAX")
  
  # Extract AE columns
  pfizer_ae <- pfizer_data[[1]]
  pfizer_bivalent_ae <- pfizer_bivalent_data[[1]]
  moderna_ae <- moderna_data[[1]]
  moderna_bivalent_ae <- moderna_bivalent_data[[1]]
  janssen_ae <- janssen_data[[1]]
  novavax_ae <- novavax_data[[1]]
  
  venn_list <- list(
    PFIZER = pfizer_ae,
    PFIZER_BIVALENT = pfizer_bivalent_ae,
    MODERNA = moderna_ae,
    MODERNA_BIVALENT = moderna_bivalent_ae,
    JANSSEN = janssen_ae,
    NOVAVAX = novavax_ae
  )
  
  custom_colors <- c(
    PFIZER = "#1f77b4",          # Blue
    PFIZER_BIVALENT = "#4b91d3", # Lighter Blue
    MODERNA = "#d62728",         # Red
    MODERNA_BIVALENT = "#ff9896",# Lighter Red
    JANSSEN = "#2ca02c",         # Green
    NOVAVAX = "#ffd700"          # Yellow
  )
  

  ####################################
  #####  ggVennDiagram
  ####################################
  # Generate the Venn diagram with custom border colors
  venn_plot <- ggVennDiagram(
    venn_list,
    label = "count",
    label_alpha = 0,
    set_color = custom_colors # Pass the custom colors here
  ) +
    scale_fill_gradient(low = "white", high = "red") # Keep the default fill gradient
  
  # # Print the plot
  # print(venn_plot)
  
  # Save ggVennDiagram plot
  ggsave(file.path(output_dir, paste0(prefix, "_ggVenn_Diagram.pdf")),
         venn_plot, width = 12, height = 8)
  
  
  ####################################
  #####  VennDetail default plot
  ####################################
  # Create other diagram using VennDetail
  six_way_venn <- VennDetail::venndetail(venn_list)
  
  # Save different plot types with specified dimensions
  pdf(file.path(output_dir, paste0(prefix, "_Venn_Diagram_Default.pdf")), 
      width = 20, height = 8)
  plot(six_way_venn)
  dev.off()
  
  
  ####################################
  #####  Upset plot - VennDetail
  ####################################
  
  # for unknown reason, PDF files are not properly created with upset plot
  upset_plot <- plot(six_way_venn, type = "upset")
  pdf(file.path(output_dir, paste0(prefix, "_Venn_Diagram_Upset_v1.pdf")), 
      width = 20, height = 8)
  print(upset_plot)
  dev.off()
  
  
  ##############################################################
  #####  Reorder custom_colors
  ##############################################################
  
  # Create a new binary_matrix
  binary_matrix <- do.call(rbind, lapply(names(venn_list), function(group) {
    items <- venn_list[[group]]
    matrix_row <- setNames(rep(0, length(unique(unlist(venn_list)))),
                           unique(unlist(venn_list)))
    matrix_row[items] <- 1
    matrix_row
  }))

  rownames(binary_matrix) <- names(venn_list)
  binary_matrix <- as.data.frame(t(binary_matrix))

  # Reorder 'custom_colors' based on the order indices
  custom_colors <- custom_colors[order(colSums(binary_matrix), decreasing = TRUE)]


  ##############################################################
  #####  Upset plot - VennDetail - with updated custom functions
  ##############################################################
  
  # for unknown reason, PDF files are not properly created with upset plot
  upset_plot_updated <- plot(six_way_venn, type="upset", mycol=custom_colors)
  pdf(file.path(output_dir, paste0(prefix, "_Venn_Diagram_Upset_v2.pdf")), 
      width = 20, height = 8)
  print(upset_plot_updated)
  dev.off()
  
  # for unknown reason, PDF files are not properly created with upset plot
  upset_plot_updated <- plot(six_way_venn, type="upset", mycol=custom_colors, 
       intersect.mycol.palette.use=TRUE)
  pdf(file.path(output_dir, paste0(prefix, "_Venn_Diagram_Upset_v3.pdf")), 
      width = 20, height = 8)
  print(upset_plot_updated)
  dev.off()

  ###################################################
  #####  Save detailed results from VennDetail object
  ###################################################
  
  venn_detail <- result(six_way_venn)
  combined_detail <- venn_detail %>%
    left_join(pfizer_data, by = c("Detail" = "AE")) %>%
    left_join(pfizer_bivalent_data, by = c("Detail" = "AE")) %>%
    left_join(moderna_data, by = c("Detail" = "AE")) %>%
    left_join(moderna_bivalent_data, by = c("Detail" = "AE")) %>%
    left_join(janssen_data, by = c("Detail" = "AE")) %>%
    left_join(novavax_data, by = c("Detail" = "AE"))
  
  write_xlsx(combined_detail, 
             file.path(output_dir, paste0(prefix, "_Combined_Output.xlsx")))
  
  
  ####################################
  #####  Upset plot - custom UpSetR
  ####################################
  

  # n_rows <- length(colnames(binary_matrix))
  # n_cols <- length(unique(venn_detail$Subset))
  # 
  # # color for each cell depends on which row we are in
  # color <- rep(custom_colors, each = n_cols) 
  # 
  # # Create mtxcol
  # mtxcol <- data.frame(x = c(1:length(color)), color = color)
  # 
  # 
  # pdf(file.path(output_dir, paste0(prefix, "_Venn_Diagram_Upset_2.pdf")), 
  #     width = 20, height = 8)
  # myupset(data = binary_matrix, 
  #         sets = colnames(binary_matrix),
  #         group.by ="degrees",
  #         point.size=5, mat_col=mtxcol,
  #         sets.bar.color= custom_colors,
  #         main.bar.color = "gray23",
  #         line.col = "blue3")
  # dev.off()
}




# Process both files
process_venn_diagrams("./excel_output/OUTPUT9.xlsx", "NonTrimmed")

process_venn_diagrams("./trimmed_excel_output/OUTPUT9.xlsx", "Trimmed")

