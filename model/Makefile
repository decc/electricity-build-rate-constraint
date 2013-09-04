libmodel.dylib: model.o
	gcc -shared -o libmodel.dylib model.o

model.o:
	gcc -Wall -fPIC -c model.c

clean:
	rm model.o
	rm libmodel.dylib
