#include <iostream>

namespace BTree {

    class Node {
    public:

        Node *left;
        Node *right;
        int val;

        Node();

        Node(int val) {
            this->val = val;
        };

    };

    void printTree(BTree::Node *&N, int level = 1, char side = '-') {

        //print @level@side: @value, e.g. 2R: 5, i.e. "2nd level right value = 5"
        std::cout << level << ":" << side << " " << N->val << "\n";

        if (N->left)    { printTree(N->left, level + 1, 'L'); }
        if (N->right)   { printTree(N->right, level + 1, 'R'); }

    }

    Node* GetTestTree() {
        
        BTree::Node *parent = new BTree::Node(1);

        parent->left = new BTree::Node(2);
        parent->right = new BTree::Node(3);

        parent->left->left = new BTree::Node(4);
        parent->left->right = new BTree::Node(5);

        parent->right->right = new BTree::Node(6);

        return parent;

    }

    void DeleteTree(Node *&N) {
        if (N->left)   { DeleteTree(N->left);   }
        if (N->right)  { DeleteTree(N->right);  }
        delete N;
    }

}