#! /bin/bash 

echo "Установка pip"
dnf install pip -y

echo "Настройка прокси для pip"
cp pip.conf /etc/pip.conf

echo "Установка библиотек"
python3 -m pip install openpyxl transliterate configparser pandas

echo "Установка завершена"
