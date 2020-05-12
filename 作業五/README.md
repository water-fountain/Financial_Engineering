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
```Option Explicit

Public Sub GetValues(ByRef t As Variant, ByRef f As Variant, ByRef Df As Variant, Optional Instn As Integer = 1)
    Dim nPoints As Integer, i As Integer
    
    With Worksheets("GraphResult")
    nPoints = Application.Count(.Range("E:E"))
    
    ReDim t(0 To nPoints - 1) As Double
    ReDim f(0 To nPoints - 1) As Double
    ReDim Df(0 To nPoints - 1) As Double
    
    For i = 0 To nPoints - 1
        t(i) = .Cells(i + 3, 5)
        f(i) = .Cells(i + 3, 6 + Instn)
        Df(i) = .Cells(i + 3, 10 + Instn)
    Next i
    End With
End Sub

Public Function HW_path(t As Variant, f As Variant, Df As Variant, r0 As Double, a As Double, sigma As Double, N As Long) As Variant
    Dim M As Integer, i As Long, j As Long
    Dim dt As Double
    M = UBound(t)
    dt = t(1) - t(0)
    ReDim theta(0 To M) As Double
    ReDim r(0 To M, 1 To N) As Double

    For j = 1 To N
        r(0, j) = r0
        For i = 0 To M - 1
            theta(i) = Df(i) + a * f(i) + sigma ^ 2 / 2 / a * (1 - Exp(-2 * a * i * dt))
            r(i + 1, j) = r(i, j) + (theta(i) - a * r(i, j)) * dt + sigma * Sqr(dt) * rGauss()
        Next i
    Next j
    
    HW_path = r
End Function

Public Function HW_path_1(t As Variant, f As Variant, Df As Variant, r0 As Double, a As Double, sigma As Double, N As Long) As Variant
    Dim M As Integer, i As Long, j As Long
    Dim dt As Double
    M = UBound(t)
    dt = t(1) - t(0)
    ReDim theta(0 To M) As Double
    ReDim r(0 To M, 1 To N) As Double

    For j = 1 To N
        r(0, j) = r0
        For i = 0 To M - 1
            theta(i) = Df(i) + a * f(i) + sigma ^ 2 / 2 / a * (1 - Exp(-2 * a * i * dt))
            r(i + 1, j) = r(i, j) + (theta(i) - a * r(i, j)) * dt + sigma * Sqr(dt) * rGauss()
            If r(i + 1, j) < 0 Then r(i + 1, j) = 0 'Convention for r < 0, take r = 0
        Next i
    Next j
    
    HW_path_1 = r
End Function

Public Function HW_path_2(t As Variant, f As Variant, Df As Variant, r0 As Double, a As Double, sigma As Double, N As Long) As Variant
    Dim M As Integer, i As Long, j As Long
    Dim dt As Double
    M = UBound(t)
    dt = t(1) - t(0)
    ReDim theta(0 To M) As Double
    ReDim r(0 To M, 1 To N) As Double

    For j = 1 To N
        r(0, j) = r0
        For i = 0 To M - 1
            theta(i) = Df(i) + a * f(i) + sigma ^ 2 / 2 / a * (1 - Exp(-2 * a * i * dt))
            r(i + 1, j) = r(i, j) + (theta(i) - a * r(i, j)) * dt + sigma * Sqr(dt) * rGauss()
            If r(i + 1, j) < 0 Then r(i + 1, j) = -r(i + 1, j) 'Convention for r < 0, take r = -r
        Next i
    Next j
    
    HW_path_2 = r
End Function

Function Intr(r As Variant, t0 As Double, t1 As Double, dt As Double) As Variant
    Dim i As Long, j As Long, N As Long
    N = UBound(r, 2)
    ReDim sum(1 To N)
    For j = 1 To N
    sum(j) = 0
        For i = Int(t0 / dt) To Int(t1 / dt) - 1
            sum(j) = sum(j) + r(i, j) * dt
        Next i
    Next j
    Intr = sum
End Function

Function Bond(Intr As Variant) As Variant
    Dim N As Long, i As Long
    N = UBound(Intr)
    ReDim Val(1 To N) As Double
    For i = 1 To N
        Val(i) = Exp(-Intr(i))
    Next i
    Bond = Val
End Function
```
