#include "../BTree.h"
#include <iostream>

void InvertTree(BTree::Node *&N);

int main() {

    BTree::Node *n = BTree::GetTestTree();

    BTree::printTree(n);
    std::cout << "=====\n";
    InvertTree(n);
    BTree::printTree(n);

    BTree::DeleteTree(n);

}

void InvertTree(BTree::Node *&N) {

    BTree::Node *oldL = NULL;
    BTree::Node *oldR = NULL;

    if (N->left)    { InvertTree(N->left);      }
    if (N->right)   { InvertTree(N->right);     }

    if (N->left)    { oldL = N->left;           }
    if (N->right)   { oldR = N->right;          }

    if (oldL)       { N->right = oldL; }        else { N->right = NULL; }
    if (oldR)       { N->left = oldR;  }        else { N->left = NULL;  }   

}