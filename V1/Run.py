from Libraries import *

if __name__=='__main__':
    n = nlpTranslator()
    #n.say("Initialising...")
    while True:
        inp = input("YOU>> ")
        n.proscessCommand(inp)
