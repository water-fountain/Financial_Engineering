先交作業

```
#==================
# model parameters
#==================
r0 <- 0.02
thetaQ <- 0.05
kappa <- 0.1
sigma <- 0.015
T0 <- 1
 
#=======================
# simulation parameters
#=======================

hedge.points <-252
simulations.total <- 40
maturity <- T0


dt <- maturity/(hedge.points+1)
timeline <- seq(0,maturity, dt)

f <- matrix(r0,(hedge.points+2),simulations.total) 

vasicek_rate <- function(r,kappa,theta,sigma,dt){
expkappadt <- exp(-kappa*dt)
vasi_vola <- (sigma^2)*(1-expkappadt^2)/(2*kappa)
result <- r*expkappadt+theta*(1-expkappadt)+sqrt(vasi_vola)*rnorm(1)
return(result)
 }

for(i in 2:(hedge.points+2)){
for(j in 1:simulations.total){
f[i,j]<- vasicek_rate(f[i-1,j],kappa,thetaQ,sigma,dt)
 }
 }
 
#==================================
# plot of interest rate simulations
#==================================

plot(timeline,f[,1], ylim=range(f,thetaQ), type="l", col="mediumorchid2")
for(j in 2:simulations.total){
lines(timeline,f[,j], col=colors()[floor(runif(1,1,657))] )
 }
abline( h = thetaQ, col = "red")
text(0, thetaQ+0.005,paste("long term interest level: theta =",thetaQ),adj=0)
title(main="Simulation of Vasicek interest rate", col.main="red", font.main=4)
```
