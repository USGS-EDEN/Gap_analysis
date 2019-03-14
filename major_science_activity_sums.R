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
emp_file <- './Gap_analysis_orig_files/OM300 Employee assignments 8-15-2017.xlsx'
emp_sheet <- 'FY18 plan 8-15-2017'

e <- read_xlsx(emp_file, emp_sheet, col_names = c("name", "title", "appt", "hours", "project_num", "task", "project_title"), col_types = c("skip", "skip", rep("text", 3), "skip", "skip", "numeric", "text", "numeric", rep("skip", 4), "text", rep("skip", 3)), skip = 1)

# Cases where title and/or appt are blank (or inconsistent)
for (i in 1:length(unique(e$name))) {
  e$title[which(e$name == unique(e$name)[i])] <- e$title[which(e$name == unique(e$name)[i] & !is.na(e$title))][1]
  e$appt[which(e$name == unique(e$name)[i])]  <- e$appt[which(e$name == unique(e$name)[i] & !is.na(e$appt))][1]
}
e$project_num[is.na(e$project_num)]     <- "Unassigned"
e$project_title[is.na(e$project_title)] <- "Unassigned"
e$task[is.na(e$task)] <- 0
e$rge <- grepl("Research", e$title)

e_sum <- group_by(e, name, title, appt, rge)
e_sum <- summarize(e_sum, thours = sum(hours, na.rm = T))

# Input file of project numbers, account names, and funding
mon_file <- './Gap_analysis_orig_files/FY18 gross project funds.xlsx'
mon_sheet <- 'funding for WFP'

m <- read_xlsx(mon_file, mon_sheet, col_names = c("proj", "wbs", "acct", "plan_sum", "late_sum"), skip = 4)

# Fill blank project numbers; remove total rows
for (i in 1:length(m$proj)) if (is.na(m$proj[i])) m$proj[i] <- m$proj[i - 1]
m <- m[-grep("Total", m$proj), ]
# Remove zero/blank rows
m <- m[!((is.na(m$plan_sum) | m$plan_sum == 0) & (is.na(m$late_sum) | m$late_sum == 0)), ]
m$plan_sum[is.na(m$plan_sum)] <- 0
m$late_sum[is.na(m$late_sum)] <- 0

# Kludge.
m$task <- read_xlsx("./Gap_analysis_orig_files/associate_tasks_to_money.xlsx", range = cell_cols("B:B"))

# Input file of science directions for tasks
sci_file <- './Gap_analysis_orig_files/tasks_associated_major_science_directions_v3.xlsx'

s <- read_xlsx(sci_file, col_names = c("proj", "task", paste0("s", 1:14)), col_types = c("skip", "text", "numeric", "skip", rep("numeric", 14)), skip = 1)
s$task[is.na(s$task)] <- 0
s[, 3:16][is.na(s[, 3:16])] <- 0

gxp <- gxl <- grcp <- grcl <- matrix(0, length(m$task), 14)
for (i in 1:length(m$proj)) {
  j <- which(s$proj == m$proj[i] & s$task == m$task[i])
  if (length(j) == 0) {
    print("No project match")
    next
  }
  a <- substr(m$wbs[i], 1, 2)
  if (a == 'GX') {
    gxl[i, ] <- unlist(m$late_sum[i] * s[j, 3:16])
    gxp[i, ] <- unlist(m$plan_sum[i] * s[j, 3:16])
  } else if (a == 'GR' | a == 'GC') {
    grcl[i, ] <- unlist(m$late_sum[i] * s[j, 3:16])
    grcp[i, ] <- unlist(m$plan_sum[i] * s[j, 3:16])
  } else print("Didn't find code")
}

png("./output/major_science_activity_sums.png", width = 1500, height = 1125)
par(mfrow = c(2, 1))
dat_p <- matrix(data = c(colSums(gxp), colSums(grcp)) / 1000, nrow = 2, byrow = T)
barplot(dat_p, col = c("darkgreen", "cyan"), main = "Sum of Original Gross Estimate FY18 B+ Plan", ylab = "$1K", ylim = c(0, 1500), cex.axis = 1.5, cex.names = 1.5)
legend("topright", c("GX", "GR & GC"), fill = c("darkgreen", "cyan"))
dat_l <- matrix(data = c(colSums(gxl), colSums(grcl)) / 1000, nrow = 2, byrow = T)
barplot(dat_l, names.arg = 1:14, col = c("darkgreen", "cyan"), main = "Sum of Plan B+ as of 05/23/2018", xlab = "Major Science Activity", ylab = "$1K", ylim = c(0, 1500), cex.axis = 1.5, cex.names = 1.5)
dev.off()

numProjTasks <- numHoursRGEPerm <- numHoursRGEPermLeave <- numHoursNonRGEPerm <- numHoursNonRGEPermLeave <- numHoursRGETerm <- numHoursRGETermLeave <- numHoursNonRGETerm <- numHoursNonRGETermLeave <- rep(0, 14)
not_done <- rep(T, length(unique(e$project_num)))
for (i in 1:14) {
  fP <- which(s[, i + 2] != 0)
  for (j in fP) {
    iT <- which(e$project_num == s$proj[j] & e$task == s$task[j])
    not_done[which(unique(e$project_num) == s$proj[j])] <- F
    numProjTasks[i] <- numProjTasks[i] + 1
    for (k in iT) {
      #print(paste(k, s$proj[j], s$task[j], e$name[k], e$rge[k], e$appt[k], e$hours[k], unlist(s[j, i + 2]), e$hours[k] * unlist(s[j, i + 2])))
      if (e$name[k] == 'Long, Joseph W' | e$name[k] == 'Stockdon, Hilary F') {
        print(paste('Employee left/leaving:', e$name[k]))
        numHoursRGEPermLeave[i] <- numHoursRGEPermLeave[i] + e$hours[k] * unlist(s[j, i + 2])
      } else if (e$name[k] == 'Khan, Nicole S' | e$name[k] == 'Locker, Stanley D') {
        print(paste('Employee left/leaving:', e$name[k]))
        numHoursRGETermLeave[i] <- numHoursRGETermLeave[i] + e$hours[k] * unlist(s[j, i + 2])
      } else if (e$name[k] == 'Barrera, Kira E' | e$name[k] == 'Brenner, Owen T') {
        print(paste('Employee left/leaving:', e$name[k]))
        numHoursNonRGETermLeave[i] <- numHoursNonRGETermLeave[i] + e$hours[k] * unlist(s[j, i + 2])
      } else if (e$rge[k] & e$appt[k] == "FTP")
        numHoursRGEPerm[i] <- numHoursRGEPerm[i] + e$hours[k] * unlist(s[j, i + 2])
      else if (!e$rge[k] & e$appt[k] == "FTP")
        numHoursNonRGEPerm[i] <- numHoursNonRGEPerm[i] + e$hours[k] * unlist(s[j, i + 2])
      else if (e$rge[k] & (e$appt[k] == "FTO" | e$appt[k] == "FEP"))
        numHoursRGETerm[i] <- numHoursRGETerm[i] + e$hours[k] * unlist(s[j, i + 2])
      else if (!e$rge[k] & (e$appt[k] == "FTO" | e$appt[k] == "FEP"))
        numHoursNonRGETerm[i] <- numHoursNonRGETerm[i] + e$hours[k] * unlist(s[j, i + 2])
      else
        print(paste('Invalid employment type for:', e$name[k], "i:", i, "proj & task:", s$proj[j], s$task[j]))
    }
  }
}
write.table(unique(e$project_title)[not_done], "./output/excluded_projects.csv", quote = F, row.names = F, col.names = F)
write.table(unique(e$project_title)[!not_done], "./output/included_projects.csv", quote = F, row.names = F, col.names = F)

tot <- sum(numHoursRGEPerm, numHoursRGEPermLeave, numHoursRGETerm, numHoursRGETermLeave, numHoursNonRGEPerm, numHoursNonRGEPermLeave, numHoursNonRGETerm, numHoursNonRGETermLeave)
png("./output/employee_science_activity_hours.png", width = 1500, height = 1125)
par(mfrow = c(2, 1))
dat_rge <- matrix(data = 100 * c(numHoursRGEPerm, numHoursRGEPermLeave, numHoursRGETerm, numHoursRGETermLeave) / tot, nrow = 4, byrow = T)
barplot(dat_rge, col = c("blue", "lightblue", "black", "grey"), main = "RGE", ylab = "% of Total Science Activity Hours", ylim = c(0, 10), cex.axis = 1.5, cex.names = 1.5)
legend("topright", c("Perm", "Perm/Left", "Term", "Term/Left"), fill = c("blue", "lightblue", "black", "grey"))
dat_nrge <- matrix(data = 100 * c(numHoursNonRGEPerm, numHoursNonRGEPermLeave, numHoursNonRGETerm, numHoursNonRGETermLeave) / tot, nrow = 4, byrow = T)
barplot(dat_nrge, names.arg = 1:14, col = c("blue", "lightblue", "black", "grey"), main = "Non-RGE", xlab = "Major Science Activity", ylab = "% of Total Science Activity Hours", ylim = c(0, 10), cex.axis = 1.5, cex.names = 1.5)
dev.off()
