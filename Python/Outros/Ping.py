import os

print("#" * 60)

ip_ou_host = input("Informe o IP ou Host:")

print("-" * 60)

os.system('ping -n 6 {}'.format(ip_ou_host))

print("-" * 60)