
#include "ExcelDriverlite.hpp"
#include "Utilities.hpp"
#include "EuropeanOption.hpp"

#include <boost/numeric/ublas/matrix.hpp>
#include <boost/numeric/ublas/io.hpp>
#include <boost/numeric/ublas/matrix_proxy.hpp>

#include <string>
#include <vector>
#include <list>
using namespace std;

//creates a linearly spaced vector from min to max with spacing = mesh_size

vector<double> CreateVec(double min, double max, double mesh_size){
    vector<double> OptionMesh;
    for (double i = min; i <= max ; i += mesh_size){
        OptionMesh.push_back(i);
    }
    return OptionMesh;
}

void PrintVec(vector<double> OptionMesh){
    for (unsigned int i = 0; i < OptionMesh.size() - 1 ; i++){
        cout << OptionMesh[i] << ", ";
    }
    cout << OptionMesh[OptionMesh.size()-1] << endl;
}



int main()
{
	EuropeanOption call_b1;
	call_b1.T = 0.25;
    call_b1.K = 65;
    call_b1.sig = 0.30;
    call_b1.r = 0.08;
    call_b1.S = 60;
    call_b1.optype = "C";

	vector<double> SpotVec = CreateVec(50,150,10);
	vector<double> SpotPriceVec;
	typedef boost::numeric::ublas::matrix<double> NumericMatrix;

	for (unsigned int i = 0; i < SpotVec.size(); i++){
		call_b1.S = SpotVec[i];
		SpotPriceVec.push_back(call_b1.Price());
	}



	ExcelDriver& excel = ExcelDriver::Instance();
	excel.MakeVisible(true);

	string sheetName("Spot Price Vector");
	string rowLabel ("Call Price");
	list<string> colLabels;
	for (unsigned int i = 0; i < SpotVec.size(); i++){
		colLabels.push_back("Spot Price = " + to_string(static_cast<long double>(SpotVec[i])));
	}

		try
	{
		long row = 1; long col = 1;
		excel.AddVector<NumericMatrix>(SpotPriceVec, sheetName, rowLabel, colLabels, row, col);
	}
	catch(std::out_of_range& e)
	{
		std::cout << e.what() << '\n';
	}
	catch (...)
	{
		// Catches everything else
		std::cout << "oop";
	}
	return 0;
}
