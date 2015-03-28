Compares docx document content.
Usage: [<options>] <expectedDocx> <actualDocx>
  options:
     /wordDiff:  Consider documents equal if Word finds no revisions comparing documents (requires Word to be installed)

Exit codes:
  -1: Error
   0: Documents are equal
   1: Differences found between documents
