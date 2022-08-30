#include <iostream>
using namespace std;

int main(){
    string strInput;
    cout << "Greetings DB, C++ has returned" << endl;
    cout << "Please enter the safe word" << endl;
    cin >> strInput ;

    if(strInput == "Bananas") {
        cout << "Thank you and goodbye" << endl;
    }
    return 0;
}

