package cn.gsein.toolkit.algorithm;

import java.lang.reflect.Array;
import java.util.Collection;
import java.util.Iterator;
import java.util.NoSuchElementException;
import java.util.Queue;

/**
 * 基于链表的队列实现
 *
 * @author G.Seinfeld
 * @date 2018/6/15
 */
public class NodeQueue<E> implements Queue<E> {

    private Node first;
    private Node last;
    private int size;

    @Override
    public int size() {
        return size;
    }

    @Override
    public boolean isEmpty() {
        return first == null;
    }

    @Override
    public boolean contains(Object o) {
        for (Node node = first; node != null; node = node.next) {
            if (o.equals(node.item)) {
                return true;
            }
        }
        return false;
    }

    @Override
    public Iterator<E> iterator() {
        return new Iterator<E>() {

            Node current = first;

            @Override
            public boolean hasNext() {
                return current != null;
            }

            @Override
            public E next() {
                return current.item;
            }
        };
    }

    @Override
    public Object[] toArray() {
        Object[] a = new Object[size];
        int k = 0;
        for (Node p = first; p != null; p = p.next) {
            a[k++] = p.item;
        }
        return a;
    }

    @SuppressWarnings("unchecked")
    @Override
    public <T> T[] toArray(T[] a) {
        if (a.length < size) {
            a = (T[]) Array.newInstance(a.getClass().getComponentType(), size);
        }
        int k = 0;
        for (Node p = first; p != null; p = p.next) {
            a[k++] = (T) p.item;
        }
        if (a.length > k) {
            a[k] = null;
        }
        return a;
    }

    @Override
    public boolean add(E e) {
        return offer(e);
    }

    @Override
    public boolean remove(Object o) {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean containsAll(Collection<?> c) {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean addAll(Collection<? extends E> c) {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean removeAll(Collection<?> c) {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean retainAll(Collection<?> c) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void clear() {
        first = null;
        last = null;
        size = 0;
    }

    @Override
    public boolean offer(E e) {
        Node oldLast = last;
        last = new Node();
        last.item = e;
        last.next = null;
        if (isEmpty()) {
            first = last;
        } else {
            oldLast.next = last;
        }
        size++;
        return true;
    }

    @Override
    public E remove() {
        if (size == 0) {
            throw new NoSuchElementException();
        }
        E item = first.item;
        first = first.next;
        if (isEmpty()) {
            last = null;
        }
        size--;
        return item;
    }

    @Override
    public E poll() {
        if (size == 0) {
            return null;
        }
        E item = first.item;
        first = first.next;
        if (isEmpty()) {
            last = null;
        }
        size--;
        return item;
    }

    @Override
    public E element() {
        return null;
    }

    @Override
    public E peek() {
        return first.item;
    }

    private class Node {
        Node next;
        E item;
    }
}
