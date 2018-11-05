library(magrittr)

combinat::permn(c(0,0,1,1)) %>%
  unique %T>%
  {set.seed(20181001)} %>%
  sample(
    size = 500,
    replace = T
  ) %>% unlist %>%
  write(
    "Projects/Replication_20181001/VB/BlockRandomisation.txt",
    ncolumns = 4
    )
  
  
