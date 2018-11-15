library(magrittr)
# setwd("\\\\sn00.zdv.uni-tuebingen.de/siive01/Documents/Projects/Replication_20181001/VB")

combinat::permn(c(0,0,1,1)) %>%
  unique %T>%
  {set.seed(20181001)} %>%
  sample(
    size = 500,
    replace = T
  ) %>% unlist %>%
  cat(
    file = "BlockRandomisation.txt"
    )