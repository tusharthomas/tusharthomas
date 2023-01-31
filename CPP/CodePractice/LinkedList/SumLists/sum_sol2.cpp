#include "../list.h"
#include <iostream>
#include <math.h>

ListNode *SumLists(ListNode *&L1, ListNode *&L2);

int main() {

    ListNode *L1 = GetTestList();
    ListNode *L2 = GetTestList();

    ListNode *L3 = SumLists(L1, L2);
    ListNode *X = L3;

    while (X) {
        std::cout << X->val << "\n";
        X = X->next;
    }

    DeleteList(L1);
    DeleteList(L2);
    DeleteList(L3);

    return 0;

}

ListNode *SumLists(ListNode *&L1, ListNode *&L2) {

    ListNode *head = new ListNode();
    ListNode *L3 = head;
    ListNode *ptr1 = L1;
    ListNode *ptr2 = L2;
    int carry = 0;
    int val = 0;

    while (ptr1 || ptr2 || carry > 0) {

        val = (ptr1? ptr1->val : 0) + (ptr2? ptr2->val : 0) + carry;
        
        if (val > 9) {
            carry = 1;
            val -= 10;
        } else {
            carry = 0;
        }

        L3->next = new ListNode(val);
        L3 = L3->next;
        if (ptr1) { ptr1 = ptr1->next? ptr1->next : NULL; }
        if (ptr2) { ptr2 = ptr2->next? ptr2->next : NULL; }
    
    }

    return head->next;

}