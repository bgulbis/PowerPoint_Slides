library(RDCOMClient)

# Create a PowerPoint application
pp <- COMCreate("Powerpoint.Application", existing=FALSE)

layouts <- pp[["SmartArtLayouts"]]
# layouts[["Count"]]
# layouts$Item(2)[["Name"]]

# Hierarchy
test <- layouts$Item(102)

name <- list()
for (i in 1:layouts[["Count"]]) {
    name[i] <- layouts$Item(i)[["Name"]]
}
# layouts(1)$Name


sa.styles <- pp[["SmartArtQuickStyles"]]
style <- list()
for (i in 1:sa.styles[["Count"]]) {
    style[i] <- sa.styles$Item(i)[["Name"]]
}

sa.colors <- pp[["SmartArtColors"]]
color <- list()
for (i in 1:sa.colors[["Count"]]) {
    color[i] <- sa.colors$Item(i)[["Name"]]
}


# Open PowerPoint
pp[["Visible"]] = TRUE

# Source the enumerated constants used by Microsoft for various parameters
source("mso.txt")

# Add a new presentation
# presentation <- pp[["Presentations"]]$Add()
# To open an existing presentation
presentation <- pp[["Presentations"]]$Open("uhcop_poster_template.pptx")
# Open(FileName As String, [ReadOnly As MsoTriState], [Untitled As MsoTriState], [WithWindow As MsoTriState = msoTrue])

# multiplication factor to convert from inches to pixels
convert.inch <- 72

# To adjust slide sizes
# presentation[["PageSetup"]][["SlideWidth"]] <- 12*convert.inch

# Add new slide
# Add(Index As Long, pCustomLayout As CustomLayout)
# slide1 <- presentation[["Slides"]]$Add(1,ms$ppLayoutBlank)

# Modify existing slide
slide1 <- presentation[["Slides"]]$Item(1)


cnt <- slide1[["Shapes"]][["Count"]]
shape <- matrix(nrow=cnt, ncol=3)
for (i in 1:cnt) {
    shape[i,1] <- slide1[["Shapes"]]$Item(i)[["Type"]]
    shape[i,2] <- slide1[["Shapes"]]$Item(i)[["Name"]]
    shape[i,3] <- slide1[["Shapes"]]$Item(i)[["Title"]]
}

# poster title
txtbx4 <- slide1[["Shapes"]]$Item(4)[["TextFrame2"]][["TextRange"]]
txtbx4[["Text"]] <- "Evaluation of Bivalirudin's Effect on the International Normalized Ratio (INR) to Determine an Appropriate Strategy for Transitioning to Warfarin"

# authors
txtbx2 <- slide1[["Shapes"]]$Item(2)[["TextFrame2"]][["TextRange"]]
txtbx2[["Text"]] <- "Andrea Fetea1, Brian E. Gulbis2, Andrea C. Hall2\r1University of Houston College of Pharmacy, 2Memorial Hermann-Texas Medical Center; Houston, TX"

# loop through all the text and make all 2's superscript
make.super <- "2"
num.super <- str_count(txtbx2[["Text"]], make.super)
txt.start <- 0

for (i in 1:num.super) {
    tmp <- txtbx2$Find(make.super, txt.start)
    txt.start <- tmp[["Start"]]
    font <- tmp[["Font"]]
    font[["Superscript"]] <- ms$msoTrue
}

# section headers
hdr1 <- slide1[["Shapes"]]$Item(7)
hdr1[["Left"]] <- 9.5*convert.inch
hdr1[["Top"]] <- 3.3*convert.inch
hdr1[["Height"]] <- 0.5*convert.inch
hdr1[["Width"]] <- 8.5*convert.inch
hdr.fill <- hdr1[["Fill"]]

# copy and paste an existing text box
hdr1$Copy()
hdr2 <- slide1[["Shapes"]]$Paste()
hdr2[["Top"]] <- 4.3*convert.inch
hdr.txt <- hdr2[["TextFrame2"]][["TextRange"]]
hdr.txt[["Text"]] <- "My Header"


# slide1 <- presentation[["Slides"]]$Add(1,ms$ppLayoutChart)
# slide2 <- presentation[["Slides"]]$Add(2,ms$ppLayoutObject)
# slide3 <- presentation[["Slides"]]$Add(3,ms$ppLayoutOrgchart)

shp1 <- slide1[["Shapes"]]$AddShape(ms$msoShape12pointStar,20,20,100,100)
shp2 <- slide1[["Shapes"]]$AddTextBox(ms$msoTextOrientationHorizontal, 50,50,100,100)
shp3 <- slide1[["Shapes"]]$AddSmartArt(test,9*convert.inch,6*convert.inch,5.5*convert.inch, 1*convert.inch)
shp.style <- shp3[["SmartArt"]]
shp.style[["QuickStyle"]] <- sa.styles$Item(3)
shp.style[["Color"]] <- sa.colors$Item(15)

# shp1_tr <- shp3[["SmartArt"]][["Nodes"]]$Item(1)[["TextFrame2"]][["TextRange"]]
# shp1_tr[["Text"]] <- "ONE"

nodes <- shp3[["SmartArt"]][["AllNodes"]][["Count"]]

for (i in 1:nodes) {
    tmp <- shp3[["SmartArt"]][["AllNodes"]]$Item(i)[["TextFrame2"]][["TextRange"]]
    tmp[["Text"]] <- paste("Sample",i)
    # tmp <- shp3[["SmartArt"]][["AllNodes"]]$Item(i)$Delete()
}
tmp <- shp3[["SmartArt"]][["AllNodes"]]$Item(6)$Delete()
tmp <- shp3[["SmartArt"]][["AllNodes"]]$Item(5)$Delete()
nodes <- shp3[["SmartArt"]][["AllNodes"]][["Count"]]

style <- shp3[["SmartArt"]][["QuickStyle"]]
style[["Id"]]

slide_width <- presentation[["PageSetup"]]$SlideWidth()
slide_height <- presentation[["PageSetup"]]$SlideHeight()

# Let's also create the rgb funtion
pp_rgb <- function(r,g,b) {
    return( r + g*256 + b*256^2)
}

# Finally, save the file in the working directory
presentation$SaveAs(paste0(getwd(),"/PowerPoint_RDCOMClient_Test.pptx"))
