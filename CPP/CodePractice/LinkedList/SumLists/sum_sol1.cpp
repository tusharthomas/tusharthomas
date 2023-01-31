/*
running into type conversion issues with very very large inputs, eg 10^30 + 1
abandoned solution
*/

#include "../list.h"
#include <iostream>
#include <math.h>

float ListToLong(ListNode *&L);
ListNode *LongToList(float num);

int main() {
    float num1 = (float) pow(10, 30) + 1;
    ListNode *L1 = LongToList(num1);
    ListNode *L2 = LongToList(465);
    float sum = ListToLong(L1) + ListToLong(L2);
    //std::cout << (char) sum << "\n";
    DeleteList(L1);
    DeleteList(L2);
}

float ListToLong(ListNode *&L) {
    int i = 0;
    float num = 0;
    while (L) {
        num += (L->val * (float) pow(10, i));
        std::cout << "num: " << num << "\n";
        i++;
        L = L->next;
    }
    return num;
    std::cout << "reached end of ListToLong\n";
}

ListNode *LongToList(float num) {
    
    float val; 
    int i = 1;
    ListNode *parent = new ListNode();
    ListNode *L = parent;

    std::cout << "num: " << num << "\n";

    while (num > 0) {
        val = (long long) num % (long long) pow(10, i);
        std::cout << "val: " << val << "\n";
        L->val = val / pow(10, i - 1);
        num -= val;
        i++;
        if (num > 0) { 
            L->next = new ListNode();
            L = L->next;
        }
    }

    return parent;

}