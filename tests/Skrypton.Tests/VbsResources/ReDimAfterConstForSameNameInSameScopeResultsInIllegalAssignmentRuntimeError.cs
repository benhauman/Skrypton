
        public void test()
        {
            const Int16 a = (Int16)1;

            _.NEWARRAY(new object[] { (Int16)1 });
            _.RAISEERROR(new IllegalAssignmentException("'a'"));
        }
