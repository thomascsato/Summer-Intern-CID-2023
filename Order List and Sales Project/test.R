abc <- ts(5 + matrix(sort(rnorm(1600)), ncol = 16, nrow = 100))
sex <- rep(c("female", "male"), each = 8)
state <- rep(c("NSW", "VIC", "QLD", "SA", "WA", "NT", "ACT", "TAS"), 2)
gc <- rbind(sex, state)  # a matrix consists of strings.
gn <- rbind(rep(1:2, each = 8), rep(1:8, 2))  # a numerical matrix
rownames(gc) <- rownames(gn) <- c("Sex", "State")
x <- gts(abc, groups = gc)
y <- gts(abc, groups = gn)

x
y
