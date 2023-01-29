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

    Node* GetTestNode() {
        return new BTree::Node(1);
    }

}