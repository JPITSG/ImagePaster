# ImagePaster Makefile
# Cross-compile for Windows using MinGW-w64

CC = x86_64-w64-mingw32-gcc
WINDRES = x86_64-w64-mingw32-windres

TARGET = ImagePaster.exe
RELEASE_DIR = release

OBJ = main.o resources.o

CFLAGS = -O2 -mwindows -I.
LDFLAGS = -mwindows
LIBS = -lshell32 -luser32 -lgdi32 -ladvapi32 -lcomctl32 -lole32 -lgdiplus

.PHONY: all clean assets

all: $(RELEASE_DIR)/$(TARGET)

$(RELEASE_DIR)/$(TARGET): $(OBJ)
	@echo "Linking executable..."
	@mkdir -p $(RELEASE_DIR)
	$(CC) -o $@ $(OBJ) $(LDFLAGS) $(LIBS)
	@rm -f $(OBJ)
	@echo "Build complete: $(RELEASE_DIR)/$(TARGET)"

main.o: main.c resource.h
	@echo "Compiling main.c..."
	$(CC) -c $< -o $@ $(CFLAGS)

resources.o: resources.rc resource.h assets/icon.ico assets/dist/index.html assets/WebView2Loader.dll
	@echo "Compiling resources..."
	$(WINDRES) $< -o $@

assets/dist/index.html: assets/package.json $(wildcard assets/src/*.tsx assets/src/**/*.tsx assets/src/**/*.ts)
	@echo "Building frontend assets..."
	cd assets && npm install && npm run build

assets: assets/dist/index.html

clean:
	rm -f $(OBJ)
	rm -rf $(RELEASE_DIR)
	rm -rf assets/dist assets/node_modules
