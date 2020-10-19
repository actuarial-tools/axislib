import struct

class BinaryReaderEOFException(Exception):
    def __init__(self):
        pass
    def __str__(self):
        return 'Not enough bytes in file to satisfy read request'

class BinaryReader:
    # Map well-known type names into struct format characters.
    typeNames = {
        'int8'   :'b',
        'uint8'  :'B',
        'int16'  :'h',
        'uint16' :'H',
        'int32'  :'i',
        'uint32' :'I',
        'int64'  :'q',
        'uint64' :'Q',
        'float'  :'f',
        'double' :'d',
        'char'   :'s'}

    def __init__(self, fileName):
        self.file = open(fileName, 'rb')
        
    def read(self, typeName):
        typeFormat = BinaryReader.typeNames[typeName.lower()]
        typeSize = struct.calcsize(typeFormat)
        value = self.file.read(typeSize)
        if typeSize != len(value):
            raise BinaryReaderEOFException
        return struct.unpack(typeFormat, value)[0]
    
    def __del__(self):
        self.file.close()

if __name__ == '__main__':
  binaryReader = BinaryReader('secret.bin')
  try:
      packet_id = binaryReader.read('uint8')
      timestamp = binaryReader.read('uint64')
      streamlen = binaryReader.read('uint32')
      streamcontent = []
      while codelength > 0:
          streamcontent.append(binaryReader.read('uint8'))
          streamlen = streamlen- 1
  except BinaryReaderEOFException:
      # One of our attempts to read a field went beyond the end of the file.
      print "Error: File seems to be corrupted."
