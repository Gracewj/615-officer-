library(officer)
library(magrittr)
library(tidyverse)
library(readxl)

##  pic a template you like
pres1 <- read_pptx("drop1.pptx") 


##  get the layout
layout_summary(pres1)

master <- "Droplet"

## 

layout_properties(x = pres1, layout = "Title Slide", master = master )

## title slide


## make a slide

pres1 %<>%  add_slide(layout = "Title Slide", master = master) %>% 
  ph_with_text(type = "ctrTitle", str = "Russel 2000 Index") %>% 
  ph_with_text(type =  "dt", str = format(Sys.Date())) %>% 
  ph_with_text(type = "sldNum", str="Slide 1") %>% 
  ph_with_text(type = "ftr", str = "Key Investment Indices")
  
  
layout_properties(x = pres1, layout = "Title and Caption", master = master )

pres1 %<>%  add_slide(layout = "Title and Caption", master = master  ) %>% 
  ph_with_text(type = "title", str = "Small Cap Index") %>% 
  ph_with_ul(type = "body", index = 1, 
               str_list = c("The Russell 2000 Index is a small-cap stock market index ",
                           "It indexes the bottom 2,000 stocks in the Russell 3000 Index "),
               level_list = c(1,1))
  

layout_properties(x = pres1, layout = "Picture with Caption", master = master )


pres1 %<>% add_slide(layout = "Picture with Caption", master = master ) %>% 
  ph_with_text(type = "title", str = "As George sees it") %>% 
  ph_with_img(type = "pic", src = "george.jpg", height = 2)



rut <- read_csv(file = "rut.csv")
ggrut <- ggplot(data=rut, aes(x=Date, y = Close)) + geom_line()

layout_properties(x = pres1, layout = "Title and Content", master = master )


pres1 %<>% add_slide(layout="Title and Content",master=master) %>%
  ph_with_text(type="title", str="Russel 2000 plot") %>%
  ph_with_gg(type="body",value=ggrut) 


print(pres1, target = "russel2000.pptx") 
