#include "MLP.h"
#include <iostream>

int main(){
    srand(time(NULL));
    rand();

    cout << "------------------- Logic Gate Example --------------------"<< endl;
    Perceptron *p = new Perceptron(2);
    p-> set_weights({51,51,-10}); //AND Approximation
    cout << "Gate:" << endl;
    cout << "0;0 :: " << p-> run({0,0}) <<endl;
    cout << "0;1 :: " << p-> run({0,1}) <<endl;
    cout << "1;0 :: " << p-> run({1,0}) <<endl;
    cout << "1;1 :: " << p-> run({1,1}) <<endl;

    return 0;
};

double frand(){

}

Perceptron::Perceptron(int inputs, double bias){
    this-> bias = bias;
    weights.resize(inputs+1);
    generate(weights.begin(), weights.end(),frand); // STL function
}

double Perceptron::run(vector<double> x){
    x.push_back(bias);
    double sum = inner_product(x.begin(), x.end(), weights.begin(), (double)0.0);
    return (sigmoid(sum));
};

void Perceptron::set_weights(vector<double> w_init){
     weights = w_init;
};

double Perceptron::sigmoid(double x){
    return 1/(1+exp(-1*x));
    // return x/(1+abs(x)); // fast sigmoid
};

MultiLayerPerception::MultiLayerPerception(vector<int> kayers, double bias, double eta){
    this->layers = layers;
    this-> bias = bias; //default = 1.0 standard practise
    this->eta = eta; //default 0.5

    for (int i = 0 ; i < layers.size(); i++){
        values.push_back(vector<double>(layers[i],0.0));
        network.push_back(vector<Perceptron>());
        if (i > 0)
            for (int j = 0; j<layers[i]; j++){
                network[i].push_back(Perceptron(layers[i-1], bias));
            }
    };
};

void MultiLayerPerceptron::set_weights(vector<vector<vector<double>>> w_init){

};

void MultiLayerPerceptron::print_weights(){
    cout << endl;
    for (int i = 1; i < network.size(); i++){
        for (int j = 0; j < layers[i]; j++){
            cout << "Layer" << i+1 << "Neuron" << j << ":" ;
            for (auto &it: network[i][j].weights)
                cout << it << "   ";
                cout << endl;
        }
    }

};