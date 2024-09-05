import random
import string


def SommeChiffres(n):
    s=0 
    while n > 0:
        s += n % 10 
        n //= 10 
    return(s)

def MulChiffres(a):
    b=1
    while a > 0:
        b *= a % 10
        a //= 10
    return(b)

def generer_string(n,m):
    taille=random.randint(n,m)
    alpha=string.ascii_lowercase
    chaine=[]
    for i in range(taille):
        chaine.append(random.choice(alpha))
    return "".join(chaine)             


def pgcd(a,b):
    if(b == 0):
        return a
    else:
        return pgcd(b, a % b)

def calcul(a,b,c):
    return a+b*c

def nbOccurences(c,chaine):
    i=0
    for j in chaine:
        if j==c:
            i=i+1
    return (i)

def listeAlphabet():
    alphabet =[]
    for i in range (26) :
        alphabet.append (chr(ord ('a')+ i ))
    return alphabet

def crypterLettre(l, cle):
    l1=(ord(l)-ord("a"))+cle
    alphabet=listeAlphabet()
    if l1>26:
        return alphabet[l1-26]
    else:
        return alphabet[l1]

def longueur_chaine(entier, chaine):
    chaine_concatenee = chaine + chaine
    
    resultat = len(chaine_concatenee) * entier
    
    return resultat