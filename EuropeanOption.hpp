//
//  EuropeanOption.hpp
//  Ex 1
//
//  Created by Ronny on 27/06/21.
//

#ifndef EuropeanOption_hpp
#define EuropeanOption_hpp

#include <iostream>
#include <string>
#include <math.h>
#include <boost/math/distributions/normal.hpp>
using namespace std;

class EuropeanOption
{
private:
    void init(); //initialise option attributes
    
    double CallPrice() const;
    double PutPrice() const;
    
public: //making option attributes public for convenience
    double T; //Expiry type of option
    double sig; //volatility
    double r; //risk free rate
    double S; //price of underlying
    double K; //strike price
    string underlying; //Underlying type
    string optype; //Option type : 'C' - Call Option, 'P' - Put Option
    

    EuropeanOption(); //default constructor
    EuropeanOption(const EuropeanOption& source); //copy constructor
    EuropeanOption(const string optiontype); //create based on option type
    virtual ~EuropeanOption(); //destructor

    EuropeanOption& operator = (const EuropeanOption& source); //overloaded = operator for EuropeanOption class
    
    double Price() const; //calculate price of option
    double ParityPrice() const; //calculate price of corresponding call/put using put call parity
    EuropeanOption toggle(); //toggle between option type C and P;
    friend bool CheckParity(const EuropeanOption& call, const EuropeanOption& put);//checks whether call and put options in arguments fulfill put call parity
       
    
};



#endif /* EuropeanOption_hpp */
