import nltk as nlp
from nltk import pos_tag as tag
import sys
from win32com.client import Dispatch
from nltk.stem.wordnet import WordNetLemmatizer
import random
import webbrowser
import requests
import bs4
import re
import string


tagsABCDEFG = '''CC	coordinating conjunction
CD	cardinal digit
DT	determiner
EX	existential there (like: "there is" ... think of it like "there exists")
FW	foreign word
IN	preposition/subordinating conjunction
JJ	adjective	'big'
JJR	adjective, comparative	'bigger'
JJS	adjective, superlative	'biggest'
LS	list marker	1)
MD	modal	could, will
NN	noun, singular 'desk'
NNS	noun plural	'desks'
NNP	proper noun, singular	'Harrison'
NNPS	proper noun, plural	'Americans'
PDT	predeterminer	'all the kids'
POS	possessive ending	parent\'s
PRP	personal pronoun	I, he, she
PRP$	possessive pronoun	my, his, hers
RB	adverb	very, silently,
RBR	adverb, comparative	better
RBS	adverb, superlative	best
RP	particle	give up
TO	to	go 'to' the store.
UH	interjection	errrrrrrrm
VB	verb, base form	take
VBD	verb, past tense	took
VBG	verb, gerund/present participle	taking
VBN	verb, past participle	taken
VBP	verb, sing. present, non-3d	take
VBZ	verb, 3rd person sing. present	takes
WDT	wh-determiner	which
WP	wh-pronoun	who, what
WP$	possessive wh-pronoun	whose
WRB	wh-abverb	where, when'''




class nlpTranslator:
    def __init__(self):
        self.continualSubject = ""
        self.quitWords = ["quit", "goodbye", "bye", "exit", "terminate", "cya", "see ya", "close"]
        self.stripPunc = str.maketrans('','','!\"\\£$%^&*()_+=-[]}{\'@~#:;.,<>?/')
        self.verbTypes = ["VB", "VBG", "VBD", "VBN", "VBP", "VBZ"]
        self.nounTypes = ["NN", "NNS", "NNP", "NNPS", "RP"]
        self.questionWords = ["WDT", "WP", "WP$", "WRB"]
        self.pronounTypes = ["PRP","PRP$"]
        self.overridePronouns = ["i", "myself", "me", "you", "he", "she", "it", "they", "we"]
        self.overrideVerbs = ["born", "die", "eat", "run", "walk"]
        self.adverbTypes = ["RB", "RBR", "RBS"]
        self.adjectiveTypes = ["JJ", "JJR", "JJS"]
        self.negatives = ["no", "not", "*nt"]
        self.yesnoWords = ["yes", "no", "ye", "yeah", "nah", "ok"]
        self.greetings = ["hello", "hi", "sup", "wasup", "hey", "morning", "yo", "wassup"]
        self.sentenceLengtheners = ["DT", "EX", "TO", "IN"]
        self.multipliers = ["CD"]
        self.speak = Dispatch("SAPI.SpVoice")
        self.lemmatizer = WordNetLemmatizer()
        self.wikiFactoriser = wikiFacts()
        self.googleFactoriser = googleFacts()
        self.sentenceObjects = ["subject", "verb", "object", "greeting", "adverb", "modal", "user", "question", "adjective", "answer"]
        self.statements = [
        "pronoun$user noun$subject verb$verb adjective$adjective",
        "pronoun$user noun$subject verb$verb noun$object",
         "noun$subject verb$verb adjective$adjective",
         "noun$subject verb$verb noun$object",
         "noun$subject verb$verb pronoun$user noun$object",
         "greeting$greeting",
         "noun$subject adverb$adverb verb$verb noun$object",
         "verb$verb noun$object",
         "noun$subject verb$verb",
         "noun$subject verb$verb noun$object",
         "noun$subject verb$verb ponoun$user noun$object",
         "verb$verb pronoun$subject noun$object",
         "yesno$answer",
         "noun$subject verb$irrelevant verb$verb noun$object",
         "noun$subject verb$irrelevant verb$verb adjective$adjective",
         "noun$answer"
         ]
        self.questions = [
         "modal$modal noun$subject verb$verb noun$user noun$object",
         "verb$verb noun$subject noun$object",
         "question$question modal$modal noun$subject verb$verb noun$object",
         "question$question verb$verb noun$subject noun$object",
         "question$question verb$verb pronoun$subject noun$object",
         "question$question verb$verb noun$subject",
         "question$question verb$verb adjective$subject",
         "question$question verb$verb adjective$adjective noun$subject",
         "verb$question noun$subject verb$verb noun$object",
         "modal$modal noun$subject verb$verb noun$object",
         "question$question",
         "question$question verb$modal noun$subject verb$verb",
         "question$question verb$modal adverb$subject verb$verb",
         "question$question verb$modal adjective$subject verb$verb",
         "question$question noun$object noun$subject"
         ]

    def normalizeVerb(self, verb):
        return self.lemmatizer.lemmatize(verb, 'v')


    def normalizeAllVerbs(self, sentence, types):
        newSent = []
        for i in range(len(sentence)):
            if types[i] == "verb":
                nml = self.normalizeVerb(sentence[i][0])
                newSent.append((nml,sentence[i][1]))
            else:
                newSent.append(sentence[i])
        return newSent

    def mergeNouns(self, sentence, simpleTypes):
        nsent = []
        nsimp = []
        offset = 0
        for i in range(len(sentence)):
            i = i + offset
            if (i < len(sentence) - 1):
                if (simpleTypes[i] == "noun") and (simpleTypes[i+1] == "noun"):
                    nsent.append((sentence[i][0] + " " +sentence[i+1][0], sentence[i][1]))
                    nsimp.append(simpleTypes[i])
                    offset = offset + 1
                else:
                    nsent.append(sentence[i])
                    nsimp.append(simpleTypes[i])
            elif (i < len(sentence)):
                nsent.append(sentence[i])
                nsimp.append(simpleTypes[i])
        return (nsent, nsimp)

    def mergeHowMany(self, sentence, simpleTypes):
        nsent = []
        nsimp = []
        offset = 0
        for i in range(len(sentence)):
            i = i + offset
            if (i < len(sentence) - 1):
                if (sentence[i][0] == "how") and (sentence[i+1][0] == "many"):
                    nsent.append((sentence[i][0] + " " +sentence[i+1][0], sentence[i][1]))
                    nsimp.append(simpleTypes[i])
                    offset = offset + 1
                else:
                    nsent.append(sentence[i])
                    nsimp.append(simpleTypes[i])
            elif (i < len(sentence)):
                nsent.append(sentence[i])
                nsimp.append(simpleTypes[i])
        return (nsent, nsimp)

    def formatSentence(self, sentence, simpleTypes, sentenceStructs):
        for struct in sentenceStructs:
            split = struct.split(" ")
            if len(split) == len(sentence):
                failed = False
                out = {}
                for (i, w) in enumerate(split):
                    halves = w.split("$")
                    varName = halves[1]
                    varType = halves[0]
                    if varType == simpleTypes[i]:
                        out[varName] = sentence[i][0]
                    else:
                        failed = True
                if not failed:
                    return out
        return None

    def removeIrrelevant(self, sentence, basicTypes):
        new = []
        newTypes = []
        for i in range(len(sentence)):
            if not (basicTypes[i] == "irrelevant" or basicTypes[i] == "multiplier"):
                new.append(sentence[i])
                newTypes.append(basicTypes[i])
        return (new, newTypes)


    def getType(self, wordTuple):
        if wordTuple[0] in self.overridePronouns:
            return "noun"
        if wordTuple[0] in self.yesnoWords:
            return "yesno"
        if wordTuple[0] in self.greetings:
            return "greeting"
        if wordTuple[1] in self.verbTypes or wordTuple[0] in self.overrideVerbs:
            return "verb"
        if wordTuple[1] in self.nounTypes:
            return "noun"
        if wordTuple[1] in self.questionWords:
            return "question"
        if wordTuple[1] in self.pronounTypes:
            return "pronoun"
        if wordTuple[1] in self.adverbTypes:
            return "adverb"
        if wordTuple[1] in self.adjectiveTypes:
            return "adjective"
        if wordTuple[1] in self.multipliers:
            return "multiplier"
        if wordTuple[1] == "CC":
            return "connector"
        if wordTuple[1] == "MD":
            return "modal"
        if wordTuple[1] in self.sentenceLengtheners:
            return "irrelevant"
        else:
            return wordTuple[1]

    def isNegative(self, wordTuple):
        score = False
        for n in self.negatives:
            if "*" in n:
                n2 = n.replace("*", "")
                if n2 in wordTuple[0]:
                    score = not score
            else:
                if n == wordTuple[0]:
                    score = not score
        return score

    def makeTypes(self, tupleList):
        types = []
        for i in tupleList:
            types.append(self.getType(i))
        return types

    def makeNegatives(self, tupleList):
        negs = []
        for i in tupleList:
            negs.append(self.isNegative(i))
        return negs

    def flip(self, word):
        if word == "me":return "you"
        if word == "my":return "your"
        if word == "mine":return "yours"
        if word == "i":return "you"
        if word == "myself":return "yourself"

        if word == "you":return "me"
        if word == "your":return "my"
        if word == "yours":return "mine"
        if word == "you":return "i"
        if word == "yourself":return "myself"

        return word

    def say(self, s):
        #s = str(s.encode("utf-8"))
        try:
            sp = s.encode("ascii", errors="ignore").decode()
        except:
            sp = s
        print("COM>> %s"%sp)
        try:
            sp = sp.replace(".", ", ")
        except:
            pass
        self.speak.Speak(sp)

    def testStucture(self, sent, struc):
        try:
            sentTypes = sent.keys()
        except:
            return False
        conditions = struc
        conditionalObjects = conditions.split(",")
        for o in conditionalObjects:
            if ":" in o:
                kvp = o.split(":")
                k = kvp[0]
                v = kvp[1]
                if not k in sent.keys():
                    return False
                if not v == sent[k]:
                    return False
            else:
                k = o
                if not k in sent.keys():
                    return False
        return True

    def removeBrackets(self, s):
        try:
            s = s.encode("ascii", errors="ignore").decode("utf-8")
        except:
            s.decode("utf-8")
        rex = re.sub(r" ?\([^)]+\)", "", s)
        rex = re.sub(r"\[[^]]+\]", "", rex)
        rex = rex.replace(")", "").replace("(", "")
        return rex

    def numSentences(self, s, num):
        try:
            return re.match(r'(?:[^.:;]+[.:;]){'+str(num)+r'}', s).group()
        except:
            return s

    def question(self, ques):
        self.say(ques)
        ans = input("YOU>> ")
        if ans in ["yes", "ye", "yeah", "sure", "why not", "ok", "y"]:
            return True
        else:
            return False

    def proscessCommand(self, s):
        rawInputCache = s
        s = s.translate(self.stripPunc).lower()
        if s in self.quitWords:
            self.say("I don't want to close")
        else:
            negScore = False
            tokens = nlp.word_tokenize(s)
            for (i,t) in enumerate(tokens):
                if self.isNegative((t, "NN")):
                    tokens.remove(t)
                    negScore = not negScore
                elif t == "im":
                    tokens.remove(t)
                    tokens.insert(i, "am")
                    tokens.insert(i, "i")
            tagged = tag(tokens)
            taggedTypes = self.makeTypes(tagged)

            #self.say(tagged)
            tagged = self.normalizeAllVerbs(tagged, taggedTypes)
            merged = self.mergeNouns(tagged, taggedTypes)
            tagged = merged[0]
            taggedTypes = merged[1]
            tup = self.removeIrrelevant(tagged, taggedTypes)
            tagged = tup[0]
            taggedTypes = tup[1]
            merged = self.mergeHowMany(tagged, taggedTypes)
            tagged = merged[0]
            taggedTypes = merged[1]
            stat = self.formatSentence(tagged, taggedTypes, self.statements)
            ques = self.formatSentence(tagged, taggedTypes, self.questions)
            if len(tagged) == 0:
                self.say("you didn't write anything")
            elif tagged[0][0] == "say":
                self.say(rawInputCache.replace("say ","",1))
            elif not stat == None:
                #self.say(stat)
                ##Special commmands
                if self.testStucture(stat, "verb:be,subject:i,adjective:hungry"):
                    self.say(["Ur so lazy. Why do i have to do everything?", "Here you go", "I have found these", "Here are some restaurants near you"][random.randrange(0, 4)])
                    webbrowser.open("https://www.google.co.uk/search?client=opera&q=restaurants+near+me&sourceid=opera&ie=UTF-8&oe=UTF-8")
                elif self.testStucture(stat, "answer"):self.say(["Ok", "Sure", "Right", "I understand"][random.randrange(0, 4)])
                elif self.testStucture(stat, "subject"):
                    d=["Good for "+self.flip(stat['subject']), "OK", "Cool", "Great"][random.randrange(0, 4)]
                    print(d)
                    self.say(d)
                else:
                    s = ["Ok", "Great", "Cool"][random.randrange(0,3)]
                    print(s)
                    self.speak(s)
            elif not ques == None:
                #self.say("Questions have not been fully implemented yet")
                #self.say(ques)
                sub = ques['subject']+" " if 'subject' in ques.keys() else ""
                adj = ques['adjective']+" " if 'adjective' in ques.keys() else ""
                pro = ques['pronoun']+" " if 'pronoun' in ques.keys() else ""
                obj = ques['object']+" " if 'object' in ques.keys() else ""
                vrb = ques["verb"] if 'verb' in ques.keys() else ""
                vrb = "" if vrb == "be" else vrb
                if ques["question"] in ["what","who", "how", "when", "where", "how many"]:
                    search = adj+pro+sub+obj+ vrb
                    self.wikiFactoriser.loadPage(search)
                    exists = self.wikiFactoriser.checkExists()
                    if not exists:
                        search2 = sub
                        self.wikiFactoriser.loadPage(search2)
                        exists = self.wikiFactoriser.checkExists()
                    if (not exists) or (self.wikiFactoriser.getSummary().replace(" ", "").replace("\n", "") == ""):
                        self.say("I don't know anything about "+search +" or "+search2)
                    else:
                        if ques['question'] in ["what", "who"]:
                            self.say(self.numSentences(self.removeBrackets(self.wikiFactoriser.getSummary()), 2))
                        elif ques["question"] == "when":
                            sents = self.removeBrackets(self.wikiFactoriser.getSummary()).split(".")
                            #print(sents)
                            useful = ""
                            for s in sents:
                                #check for number
                                if any(char.isdigit() for char in s):
                                    useful = useful+s+"."
                                break
                            if useful == "":
                                self.say("I am not sure about the date, but i have found an article")
                                #self.say(tagged)
                            else:
                                self.say(useful)
                        elif ques["question"] == "how many":
                            sents = self.removeBrackets(self.wikiFactoriser.getSummary()).split(".")
                            #print(sents)
                            useful = ""
                            for s in sents:
                                #check for number
                                if any(char.isdigit() for char in s):
                                    useful = useful+s+"."
                                break
                            if useful == "":
                                self.say("I am not sure about any numbers, but i have found an article")
                                #self.say(tagged)
                            else:
                                self.say(useful)
                        else:
                            sents = self.removeBrackets(self.wikiFactoriser.getSummary()).split(".")
                            #print(sents)
                            useful = ""
                            for s in sents:
                                if ques["question"] in s:
                                    useful = useful+"."+s
                            self.say(useful)
                        learnMore = self.question("Do you want to learn more?")
                        if learnMore:
                            webbrowser.open(self.wikiFactoriser.fullURL)
                            self.say("Opening "+str(self.wikiFactoriser.getTitle()))
                else:
                    self.say("I'm not sure how to answer that")
            else:
                self.say("Im not sure what that meant")
                self.say(tagged)
            #self.say(taggedNegatives)

class wikiFacts:
    def __init__(self):
        self.baseURL = """https://en.wikipedia.org/wiki/"""
        self.spaceReplace = "_"
        self.page = None
        self.fullURL = ""
    def loadPage(self, name):
        nameSpacesReplaced = name.replace(" ", self.spaceReplace)
        url = self.baseURL + nameSpacesReplaced
        self.fullURL = url
        pageRaw = requests.get(url)
        self.page = bs4.BeautifulSoup(pageRaw.text, 'lxml')
    def checkExists(self):
        styleCorrect = self.page.find_all("b")
        for i in styleCorrect:
            #print(i)
            if "does not have an article" in str(i):
                return False
        return True
    def getTitle(self):
        return self.page.find("title").getText()
    def getSummary(self):
        styleCorrect = self.page.find_all("p")
        try:
            p = (styleCorrect[1]).getText()
            if p.count(".") < 2 and len(styleCorrect) >= 3:
                p = (styleCorrect[2]).getText()
            if p.count(".") < 2 and len(styleCorrect) >= 4:
                p = (styleCorrect[3]).getText()
            if p.count(".") < 2 and len(styleCorrect) >= 5:
                p = (styleCorrect[4]).getText()
        except:
            try:
                p = (styleCorrect[0]).getText()
            except:
                p = ""
        return p

class googleFacts:
    #Experimental and not fully working
    def __init__(self):
        self.baseURL = """https://www.google.co.uk/search?q="""
        self.spaceReplace = "+"
        self.page = None
        self.fullURL = ""
    def loadPage(self, name):
        nameSpacesReplaced = name.replace(" ", self.spaceReplace)
        url = self.baseURL + nameSpacesReplaced
        self.fullURL = url
        pageRaw = requests.get(url)
        self.page = bs4.BeautifulSoup(pageRaw.text, 'lxml')
    def checkExists(self):
        styleCorrect = self.page.find_all("span", {"class":"ILfuVd yZ8quc"})
        #for x in styleCorrect:
            #print(x)
        if styleCorrect[1].getText() == "":
            return True
        else:
            return False
    def getTitle(self):
        return self.page.find("title").getText()
    def getFact(keywords):
        pass

if __name__=='__main__':
    n = nlpTranslator()
    #n.say("Initialising...")
    while True:
        inp = input("YOU>> ")
        n.proscessCommand(inp)
