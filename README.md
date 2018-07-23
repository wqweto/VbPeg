## VbPeg
PEG parser generator for VB6

### Description

VbPeg is a simple parser generator for VB6 that can be used to build interpreters, compilers and other tools that need to match and process  complex input data.

VbPeg generates recursive-descent parsers from Parsing Expression Grammars (PEG) [[Ford 2004]](http://bford.info/pub/lang/peg.pdf) and is currently based on the original Ford syntax with some additions for 'semantic actions' as implemented by Ian Piumarta's [peg/leg](http://piumarta.com/software/peg/) project.

VbPeg is self-hosted, meaning it can produce it's own parser in src\cParser.cls from the grammar in VbPeg.peg in root.

### Sample usage

 - Generate a private VB6 class from PEG grammar in VbPeg.peg
```
    c:> VbPeg.exe VbPeg.peg -o src\cParser
```
### Command-line
```
Usage: VbPeg.exe [options] <in_file.peg>

Options:
  -o OUTFILE      write result to OUTFILE [default: stdout]
  -tree           output parse tree
  -ir             output intermediate represetation
  -set NAME=VALUE set or modify grammar settings
  -q              in quiet operation outputs only errors
  -nologo         suppress startup banner
  -allrules       output all rules (don't skip unused)
  -trace          trace in_file.peg parsing as performed by %1.exe

If no -tree/-ir is used emits VB6 code. If no -o is used writes result to console.
```
