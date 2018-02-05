## VbPeg
PEG parser generator for VB6

### Description

VbPeg is a simple parser generator for VB6 that can be used to build interpreters, compilers and other tools that need to match and process  complex input data.

VbPeg generates recursive-descent parsers from Parsing Expression Grammars (PEG) [[Ford 2004]](http://bford.info/pub/lang/peg.pdf) and is currently based on the original Ford syntax with some additions for 'semantic actions' as implemented by Ian Piumarta's [peg/leg](http://piumarta.com/software/peg/) project.

VbPeg is self-hosted, meaning it can produce it's own parser in src\cParser.cls from the grammar in VbPeg.peg in root.

### Sample usage

 - Generate a private VB6 class from PEG grammar in VbPeg.peg
```
    c:> VbPeg.exe VbPeg.peg -o src\cParser -private
```
### Command-line
```
Usage: VbPeg.exe [options] <in_file.peg>

Options:
  -o OUTFILE      write result to OUTFILE [default: stdout]
  -tree           output parse tree
  -ir             output intermediate represetation
  -public         emit public VB6 class module
  -private        emit private VB6 class module
  -module NAME    VB6 class/module name [default: OUTFILE]
  -userdata NAME  parser context UserData member data-type [default: Variant]
  -q              in quiet operation outputs only errors

If no -tree/-ir is used emits VB6 code. If no -o is used writes result to 
console. If no -public/-private is used emits standard .bas module.
```

### ToDo

 - Extend grammar based on [PEG.js](https://github.com/pegjs/pegjs)
 - Implement named captures
 - Case-insensitive matching
