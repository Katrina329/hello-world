class Node:
    def __init__(self, data, next=None):
        self.data, self.next = data, next


class LinkList:
    def __init__(self, lst):
        self.head = Node(lst[0])
        p = self.head
        for i in lst[1:]:
            node = Node(i)
            p.next = node
            p = p.next


def genLinkList():
    a = list(map(int, input().split()))
    b = list(map(int, input().split()))
    n = int(input())
    la = LinkList(a)
    lb = LinkList(b)

    p = la.head
    last = None
    while p is not None:
        last = p
        p = p.next

    p = lb.head
    for i in range(n - 1):
        p = p.next
    last.next = p

    return la, lb


a, b = genLinkList()


def findJointPoint(a, b):
    len_a = 0
    p = a.head
    while p is not None:
        len_a += 1
        p = p.next

    len_b = 0
    p = b.head
    while p is not None:
        len_b += 1
        p = p.next

    pa, pb = a.head, b.head
    if len_a > len_b:
        for _ in range(len_a - len_b):
            pa = pa.next
    else:
        for _ in range(len_b - len_a):
            pb = pb.next

    while pa is not pb:
        pa = pa.next
        pb = pb.next

    return pa.data


print(findJointPoint(a, b))
