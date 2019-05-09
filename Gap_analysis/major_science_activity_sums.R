# -------------
# Bryan McCloskey
# bmccloskey@usgs.gov
# St. Petersburg Coastal and Marine Science Center
# US Geological Survey
#
# 03/12/2019
#--------------

print("These libraries must be installed: readxl, dplyr")
# Required libraries. If not present, run:
# install.packages("readxl")
# install.packages("dplyr")
library (readxl)
library (dplyr)

# Input file of employee names, titles, appointment type, and assigned hours
emp_file <- './input/GGESMN FY19 OM300 4.3.19.xlsx'

e <- read_xlsx(emp_file, col_names = c("name", "title", "appt", "hours", "project_num", "task", "project_title"), col_types = c(rep("skip", 2), rep("text", 3), rep("skip", 3), "numeric", "text", "numeric", rep("skip", 3), "text", rep("skip", 3)), skip = 1)

e$project_num[is.na(e$project_num)]     <- "Unassigned"
e$project_title[is.na(e$project_title)] <- "Unassigned"
e$task[is.na(e$task)] <- 0
e$rge <- grepl("Research", e$title)

e_sum <- group_by(e, name, appt, rge)
e_sum <- summarize(e_sum, thours = sum(hours))

# Input file of project numbers, account names, and funding
mon_file <- './input/GGESMN FY19 CCM 640 4.3.19.xlsx'
fy_s <- regexpr('FY', mon_file)
fy <- substr(mon_file, fy_s + 2, fy_s + 3)
dt_s <- regexpr('[0-9]{1,2}.[0-9]{1,2}.[0-9]{2}', mon_file)
dt <- substr(mon_file, dt_s, dt_s + attr(dt_s, "match.length") - 1)

m <- read_xlsx(mon_file, col_names = c("proj", "sum", "task"), skip = 3)

m <- m[-dim(m)[1], ] # remove total row
m <- data.frame(proj = m[seq(3, dim(m)[1], 3), "proj"], wbs = m[seq(1, dim(m)[1], 3), 1], sum = m[seq(1, dim(m)[1], 3), "sum"], task = m[seq(1, dim(m)[1], 3), "task"])
names(m)[2] <- "wbs"
m <- m[m$sum != 0, ]

# Input file of science directions for tasks
sci_file <- './input/tasks_associated_major_science_directions_final 4.3.19.xlsx'

s1 <- read_xlsx(sci_file, n_max = 1)
s_dim <- dim(s1)[2] - 4
s_names <- names(s1)[5:dim(s1)[2]]
s <- read_xlsx(sci_file, col_names = c("proj", "task", paste0("s", 1:s_dim)), col_types = c("skip", "text", "numeric", "skip", rep("numeric", s_dim)), skip = 1)
s$task[is.na(s$task)] <- 0
s[, 3:(s_dim + 2)][is.na(s[, 3:(s_dim + 2)])] <- 0

gx <- grc <- matrix(0, length(m$task), s_dim)
for (i in 1:length(m$proj)) {
  j <- which(s$proj == m$proj[i] & s$task == m$task[i])
  if (length(j) == 0) {
    print(paste(m$proj[i], m$task[i], "No project match"))
    next
  }
  a <- substr(m$wbs[i], 1, 2)
  if (a == 'GX') {
    gx[i, ] <- unlist(m$sum[i] * s[j, 3:(s_dim + 2)])
  } else if (a == 'GR' | a == 'GC') {
    grc[i, ] <- unlist(m$sum[i] * s[j, 3:(s_dim + 2)])
  } else print("Didn't find code")
}

numProjTasks <- numHoursRGEPerm <- numHoursRGEPermLeave <- numHoursNonRGEPerm <- numHoursNonRGEPermLeave <- numHoursRGETerm <- numHoursRGETermLeave <- numHoursNonRGETerm <- numHoursNonRGETermLeave <- rep(0, s_dim)
not_done <- rep(T, length(unique(e$project_num)))
for (i in 1:s_dim) {
  fP <- which(s[, i + 2] != 0)
  for (j in fP) {
    iT <- which(e$project_num == s$proj[j] & e$task == s$task[j])
    not_done[which(unique(e$project_num) == s$proj[j])] <- F
    numProjTasks[i] <- numProjTasks[i] + 1
    for (k in iT) {
      #print(paste(k, s$proj[j], s$task[j], e$name[k], e$rge[k], e$appt[k], e$hours[k], unlist(s[j, i + 2]), e$hours[k] * unlist(s[j, i + 2])))
      if (e$name[k] == 'Plant, Nathaniel G') {
        print(paste('Employee left/leaving:', e$name[k]))
        numHoursRGEPermLeave[i] <- numHoursRGEPermLeave[i] + e$hours[k] * unlist(s[j, i + 2])
      } else if (e$rge[k] & e$appt[k] == "FTP")
        numHoursRGEPerm[i] <- numHoursRGEPerm[i] + e$hours[k] * unlist(s[j, i + 2])
      else if (!e$rge[k] & e$appt[k] == "FTP")
        numHoursNonRGEPerm[i] <- numHoursNonRGEPerm[i] + e$hours[k] * unlist(s[j, i + 2])
      else if (e$rge[k] & (e$appt[k] == "FTO" | e$appt[k] == "FEP"))
        numHoursRGETerm[i] <- numHoursRGETerm[i] + e$hours[k] * unlist(s[j, i + 2])
      else if (!e$rge[k] & (e$appt[k] == "FTO" | e$appt[k] == "FEP" | e$appt[k] == "FEN"))
        numHoursNonRGETerm[i] <- numHoursNonRGETerm[i] + e$hours[k] * unlist(s[j, i + 2])
      else
        print(paste('Invalid employment type for:', e$name[k], "i:", i, "proj & task:", s$proj[j], s$task[j]))
    }
  }
}
write.table(unique(e$project_title)[not_done], "./output/excluded_projects.csv", quote = F, row.names = F, col.names = F)
write.table(unique(e$project_title)[!not_done], "./output/included_projects.csv", quote = F, row.names = F, col.names = F)

tot <- sum(numHoursRGEPerm, numHoursRGEPermLeave, numHoursRGETerm, numHoursRGETermLeave, numHoursNonRGEPerm, numHoursNonRGEPermLeave, numHoursNonRGETerm, numHoursNonRGETermLeave)

png("./output/employee_science_activity_hours.png", width = 1000, height = 750)
par(mfrow = c(3, 1))
dat_rge <- matrix(data = 100 * c(numHoursRGEPerm, numHoursRGEPermLeave, numHoursRGETerm, numHoursRGETermLeave) / tot, nrow = 4, byrow = T)
dat_nrge <- matrix(data = 100 * c(numHoursNonRGEPerm, numHoursNonRGEPermLeave, numHoursNonRGETerm, numHoursNonRGETermLeave) / tot, nrow = 4, byrow = T)
max <- ceiling(max(colSums(cbind(dat_rge, dat_nrge))))
barplot(dat_rge, col = c("blue", "lightblue", "black", "grey"), main = "RGE", ylab = "% of Total Science Activity Hours", ylim = c(0, max), cex.axis = 1.5)
legend("topright", c("Perm", "Perm/Left", "Term", "Term/Left"), fill = c("blue", "lightblue", "black", "grey"), cex = 1.5)
barplot(dat_nrge, col = c("blue", "lightblue", "black", "grey"), main = "Non-RGE", ylab = "% of Total Science Activity Hours", ylim = c(0, max), cex.axis = 1.5)
par(xpd=T)
legend("topright", paste(1:s_dim, s_names), inset = c(0, -.24))
dat_s <- matrix(data = c(colSums(gx), colSums(grc)) / 1000, nrow = 2, byrow = T)
max <- round(max(colSums(dat_s)), digits = -2)
barplot(dat_s, names.arg = 1:s_dim, col = c("darkgreen", "cyan"), main = paste0("SPCMSC FY", fy, " Estimated Funding ", dt), xlab = "Major Science Activity", ylab = "$1K", ylim = c(0, max), cex.axis = 1.5, cex.names = 1.5)
legend("topright", c("GX", "GR & GC"), fill = c("darkgreen", "cyan"), cex = 1.5)
dev.off()

