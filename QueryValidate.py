import datetime

def QueryValidation(instrument, queryCommand, comparedVariable, inFunction):
	queryFile = open('QueryValidation ' + str(datetime.datetime.now().strftime("%y-%m-%d"))+'.txt', 'a')
	queryFile.close()
        queryFile = open('QueryValidation ' + str(datetime.datetime.now().strftime("%y-%m-%d"))+'.txt', 'a')
        queryFile.write('-----'+str(inFunction)+'----- \n')
        queryFile.close()
	Query = instrument.query(queryCommand)
	queryFile = open('QueryValidation ' + str(datetime.datetime.now().strftime("%y-%m-%d"))+'.txt', 'a')
	queryFile.write("Does " + str(Query) + " match with " + str(comparedVariable) + " ? \n")
	queryFile.close()