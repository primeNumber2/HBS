# __author__ = 'WangS15'
# pwd = "!ws861018"
# # key = "142857572814"
# # word = []
# def gene(key):
#     word = []
#     for num in range(len(pwd)):
#         digit = ord(pwd[num])
#         # print(digit)
#         digit += int(key[num])
#         # print(digit)
#         alphabet = chr(digit)
#         # print(alphabet)
#         word.append(alphabet)
#     return "".join(word)
#     # print(ord(pwd[num]))
#
# # print("".join(word))
# print(gene("142857285714"))
# # F1ak!nj0ke
#
#
# def back(word, key):
#     password = []
#     for num in range(len(word)):
#         digit = ord(word[num])
#         digit -= int(key[num])
#         alphabet = chr(digit)
#         password.append(alphabet)
#     return "".join(password)
#
# # result = back("\"WZa6]y{E[fw", "142857285714")
# # print(result)
# print()