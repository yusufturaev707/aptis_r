user = [
    'Turaev66', 'Yusuf66', 'A'
]

users = [
    ['Turaev', 'Yusuf', 'A', 'B', 'C'],
    ['Turaev1', 'Yusuf1', 'A', 'B', 'C'],
    ['Turaev2', 'Yusuf2', 'A', 'B', 'C'],
    ['Turaev6', 'Yusuf6', 'A', 'B', 'C'],
    ['Turaev4', 'Yusuf4', 'A', 'B', 'C'],
    ['Turaev5', 'Yusuf5', 'A', 'B', 'C'],
    ['Turaev6', 'Yusuf6', 'A', 'B', 'C'],
    ['Turaev7', 'Yusuf7', 'A', 'B', 'C'],
]

for us in users:
    if not (user[0] and user[1]) in us:
        print(True)