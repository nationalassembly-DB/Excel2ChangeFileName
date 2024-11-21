from natsort import natsorted

# Example strings
items = [
    '2023년도_교육위원회_국정감사_서면답변서(경상남도교육청)(2).PDF', '2023년도_교육위원회_국정감사_서면답변서(경상남도교육청).PDF'
]

# Sorting using natsort
sorted_items = natsorted(items)
print(sorted_items)
