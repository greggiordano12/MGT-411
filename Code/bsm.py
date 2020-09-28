import numpy as np
import scipy.stats as si
import sympy as sy
from sympy.stats import Normal, cdf
from sympy import init_printing
def black_scholes(S, K, T, r, sigma, option_type="call"):

    #S: spot price
    #K: strike price
    #T: time to maturity
    #r: interest rate
    #sigma: volatility of underlying asset

    d1 = (np.log(S / K) + (r + 0.5 * sigma ** 2) * T) / (sigma * np.sqrt(T))
    d2 = (np.log(S / K) + (r - 0.5 * sigma ** 2) * T) / (sigma * np.sqrt(T))

    call = (S * si.norm.cdf(d1, 0.0, 1.0) - K * np.exp(-r * T) * si.norm.cdf(d2, 0.0, 1.0))
    put = (K * np.exp(-r * T) * si.norm.cdf(-d2, 0.0, 1.0) - S * si.norm.cdf(-d1, 0.0, 1.0))
    if option_type == "call":
        return call
    else:
        return put
S = 50
K = 100
T = 1
r = .05
sigma = .25
black_scholes(S,K,T,r,sigma)
black_scholes(S,K,T,r,sigma,"put")

S = 341.58
K = 340
T = 4/365
r = .02
option_price = 3.64

def bsm_zero(S, K, T, r, sigma, option_price, option_type="call"):
    d1 = (np.log(S / K) + (r + 0.5 * sigma ** 2) * T) / (sigma * np.sqrt(T))
    d2 = (np.log(S / K) + (r - 0.5 * sigma ** 2) * T) / (sigma * np.sqrt(T))

    call = (S * si.norm.cdf(d1, 0.0, 1.0) - K * np.exp(-r * T) * si.norm.cdf(d2, 0.0, 1.0))
    put = (K * np.exp(-r * T) * si.norm.cdf(-d2, 0.0, 1.0) - S * si.norm.cdf(-d1, 0.0, 1.0))
    if option_type == "call":
        return (call - option_price)
    else:
        return (put-option_price)

def secant(S=100,K=100,T=1,r=.02,x1=0, x2=1, option_price=2, option_type = "call",E=.00001):
    n = 0; xm = 0; x0 = 0; c = 0;
    if (bsm_zero(S,K,T,r,x1,option_price,option_type) * bsm_zero(S,K,T,r,x2,option_price,option_type) < 0):
        while True:

            # calculate the intermediate value
            x0 = ((x1 * bsm_zero(S,K,T,r,x2,option_price,option_type) - x2 * bsm_zero(S,K,T,r,x1,option_price,option_type)) /
                            (bsm_zero(S,K,T,r,x2,option_price,option_type) - bsm_zero(S,K,T,r,x1,option_price,option_type)));

            # check if x0 is root of
            # equation or not
            c = bsm_zero(S,K,T,r,x1,option_price,option_type) * bsm_zero(S,K,T,r,x0,option_price,option_type);

            # update the value of interval
            x1 = x2;
            x2 = x0;

            # update number of iteration
            n += 1;

            # if x0 is the root of equation
            # then break the loop
            if (c == 0):
                break;
            xm = ((x1 * bsm_zero(S,K,T,r,x2,option_price,option_type) - x2 * bsm_zero(S,K,T,r,x1,option_price,option_type)) /
                            (bsm_zero(S,K,T,r,x2,option_price,option_type) - bsm_zero(S,K,T,r,x1,option_price,option_type)));

            if(abs(xm - x0) < E):
                break;

        print("Root of the given equation =",
                               round(x0, 6));
        print("No. of iterations = ", n);

    else:
        print("Can not find a root in ",
                   "the given inteval");


secant(S,K,T,r,option_price=option_price)
