//
//  EuropeanOption.cpp
//  Ex 1
//
//  Created by Ronny on 27/06/21.
//

#include "EuropeanOption.hpp"
#include <cstdlib>


void EuropeanOption::init(){ //initiating option attributes to default values;
    T = 1;
    r = 0.05;
    sig = 0.2;
    K = 110;
    S = 100;
    underlying = "Stock";
    optype = "C";
}


EuropeanOption::EuropeanOption(){
    init();
}

EuropeanOption::EuropeanOption(const EuropeanOption& source){
    T = source.T;
    r = source.r;
    sig = source.sig;
    K = source.K;
    S = source.S;
    underlying = source.underlying;
    optype = source.optype;
}

EuropeanOption::EuropeanOption(const string optiontype){
    init();
    optype = optiontype;
    if (optype == "c") optype = "C";
    else if (optype == "p") optype = "P";
}

EuropeanOption::~EuropeanOption(){
}

EuropeanOption& EuropeanOption::operator = (const EuropeanOption& source){
    if (this == &source) return *this;
    T = source.T;
    r = source.r;
    sig = source.sig;
    K = source.K;
    S = source.S;
    underlying = source.underlying;
    optype = source.optype;
    return *this;
}

double EuropeanOption::CallPrice() const{
    boost::math::normal_distribution<double> nfunc(0,1); //creating standard normal object with mean 0 and std dev 1;
    double d1,d2,N1,N2,price = 0;
    d1 = 0;
    if (underlying == "Stock"){
    d1 = (log(S/K) + ((r + 0.5*pow(sig,2))*T)) / (sig*sqrt(T));
    }
    else if (underlying == "Futures"){
    d1 = (log(S/K) + (0.5*pow(sig,2))*T) / (sig*sqrt(T));
    }
    d2 = d1 - (sig*sqrt(T));
    N1 = boost::math::cdf(nfunc,d1);
    N2 = boost::math::cdf(nfunc,d2);
    
    if (underlying =="Stock") price = (S*N1) - K*(exp(-r*T))*N2;
    else if (underlying == "Futures") price = (S*exp(-r*T)*N1) - K*(exp(-r*T))*N2;
    return price;
}

double EuropeanOption::PutPrice() const{
    boost::math::normal_distribution<double> nfunc(0,1); //creating standard normal object with mean 0 and std dev 1;
    double d1,d2,N1,N2,price = 0;
    d1 = 0;
    
    if (underlying == "Stock"){
    d1 = (log(S/K) + ((r + 0.5*pow(sig,2))*T)) / (sig*sqrt(T));
    }
    else if (underlying == "Futures"){
    d1 = (log(S/K) + (0.5*pow(sig,2))*T) / (sig*sqrt(T));
    }
    
    d2 = d1 - (sig*sqrt(T));
    N1 = boost::math::cdf(nfunc,-d1);
    N2 = boost::math::cdf(nfunc,-d2);
    
    if (underlying =="Stock") price =  K*(exp(-r*T))*N2 - (S*N1);
    else if (underlying == "Futures") price =  K*(exp(-r*T))*N2 - (S*exp(-r*T)*N1);
    return price;
}

double EuropeanOption::Price() const{
    if (optype == "C") return CallPrice();
    return PutPrice();
}

EuropeanOption EuropeanOption::toggle(){
    EuropeanOption temp(*this);
    if (temp.optype == "C") temp.optype = "P";
    else temp.optype = "C";
    return temp;
}

double EuropeanOption::ParityPrice() const{
    if (optype == "C") return (Price() - S + (K*exp(-r*T)));
    return (Price() + S - (K*exp(-r*T)));
}

bool CheckParity(const EuropeanOption& call, const EuropeanOption& put){
    double epsilon = 0.00001; //precision
    double call_price = call.Price();
    double put_price = put.Price();
    double r = call.r;
    double K = call.K;
    double S = call.S;
    double T = call.T;
    double p1 = call_price + K*exp(-r*T);
    double p2 = put_price + S;
    if (abs(p1 - p2) < epsilon) return true;
    else return false;
}
    
