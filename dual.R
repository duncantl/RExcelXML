cmp = 
function(what, dir = c("template", "Truth"))
{
  ff = sprintf("%s/%s", dir, what)
  structure(lapply(ff, xmlParse), names = dir)
}
