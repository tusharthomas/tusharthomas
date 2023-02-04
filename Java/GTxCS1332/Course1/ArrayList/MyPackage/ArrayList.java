package MyPackage;
import java.util.NoSuchElementException;
import java.util.prefs.BackingStoreException;

/**
 * Your implementation of an ArrayList.
 */
public class ArrayList<T> {

    /*
     * The initial capacity of the ArrayList.
     *
     * DO NOT MODIFY THIS VARIABLE!
     */
    public static final int INITIAL_CAPACITY = 9;

    /*
     * Do not add new instance variables or modify existing ones.
     */
    private T[] backingArray;
    private int size;

    /**
     * This is the constructor that constructs a new ArrayList.
     * 
     * Recall that Java does not allow for regular generic array creation,
     * so instead we cast an Object[] to a T[] to get the generic typing.
     */
    public ArrayList() {
        //DO NOT MODIFY THIS METHOD!
        backingArray = (T[]) new Object[INITIAL_CAPACITY];
    }

    /**
     * Adds the data to the front of the list.
     *
     * This add may require elements to be shifted.
     *
     * Method should run in O(n) time.
     *
     * @param data the data to add to the front of the list
     * @throws java.lang.IllegalArgumentException if data is null
     */
    public void addToFront(T data) {
        if (data == null) {
            throw new java.lang.IllegalArgumentException();
        } else {
            if (size >0) {
                if (size >= backingArray.length) { resize(); }
                shiftBack(0);
            }
            backingArray[0] = data;
            size++;
        }
    }

    /**
     * Adds the data to the back of the list.
     *
     * Method should run in amortized O(1) time.
     *
     * @param data the data to add to the back of the list
     * @throws java.lang.IllegalArgumentException if data is null
     */
    public void addToBack(T data) {
        if (data == null) {
            throw new java.lang.IllegalArgumentException();
        }
        else {
            if (size >= backingArray.length) { resize(); }
            backingArray[size] = data;
            size++;
        }
    }

    private void resize() {
        T[] newArray = (T[]) new Object[backingArray.length * 2];
        int i;
        for (i = 0; i < backingArray.length; i++) {
            newArray[i] = backingArray[i];
        }
        backingArray = newArray;
    }

    private void shiftBack(int startPos) {
        int i;
        for (i = size; i > startPos; i--) {
            backingArray[i] = backingArray[i - 1];
        }
    }

    private void shiftForward(int startPos) {
        int i;
        for (i = startPos; i < size - 1; i++ ) {
            backingArray[i] = backingArray[i + 1];
        }
    }

    public void printArrayList() {
        int i;
        for (i = 0; i <= backingArray.length - 1; i++) {
            System.out.println(backingArray[i]);
        }
    }

    /**
     * Removes and returns the first data of the list.
     *
     * Do not shrink the backing array.
     *
     * This remove may require elements to be shifted.
     *
     * Method should run in O(n) time.
     *
     * @return the data formerly located at the front of the list
     * @throws java.util.NoSuchElementException if the list is empty
     */
    public T removeFromFront() {
        T data;
        if (backingArray[0] == null) {
            throw new java.util.NoSuchElementException();
        } else {
            data = backingArray[0];
            shiftForward(0);
            size--;
            backingArray[size] = null;
            return data;
        }
    }

    /**
     * Removes and returns the last data of the list.
     *
     * Do not shrink the backing array.
     *
     * Method should run in O(1) time.
     *
     * @return the data formerly located at the back of the list
     * @throws java.util.NoSuchElementException if the list is empty
     */
    public T removeFromBack() {
        T data;
        if (backingArray[0] == null) {
            throw new java.util.NoSuchElementException();
        } else {
            size--;
            data = backingArray[size];
            backingArray[size] = null;
            return data;
        }
    }

    /**
     * Returns the backing array of the list.
     *
     * For grading purposes only. You shouldn't need to use this method since
     * you have direct access to the variable.
     *
     * @return the backing array of the list
     */
    public T[] getBackingArray() {
        // DO NOT MODIFY THIS METHOD!
        return backingArray;
    }

    /**
     * Returns the size of the list.
     *
     * For grading purposes only. You shouldn't need to use this method since
     * you have direct access to the variable.
     *
     * @return the size of the list
     */
    public int size() {
        // DO NOT MODIFY THIS METHOD!
        return size;
    }
}