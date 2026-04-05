public void test()
{
    const Int16 a = (Int16)1;

    _.NEWARRAY(new object[] { (Int16)1 });
    _.RAISEERROR(new Skrypton.RuntimeSupport.IllegalAssignmentException("'a' : The left-hand side of an assignment must be a variable, property or indexer and not <constant>"));
}