import sqlite3

def conectar():
    conn = sqlite3.connect('database.db')
    cursor =  conn.cursor()
    return cursor, conn