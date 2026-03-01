=LET(
  _Author,N("your name"),
  _Source,N("formulas/myformula.gs"),
  _Deployed_using,N("python deploy_gs.py 'Spreadsheet Name' 'SheetName!$C$14'"),
  _Date_deployed,N("deployment date"),

  /* Prime Number Calculator */

  TextJoin(", ",,
    Filter(
      Sequence(199),          /* For each number from 1 to 199, filter it based on the True/False values returned by Map */
      Map(                    /*    Map returns 199 True/False values, where True indicates a prime number.              */
        Sequence(199),        /* For x=1 to 199, use the following test to determine if x is prime:                      */
        Lambda(x,
          1=Countif(          /*    return True if Mod(x,y)=0 exactly once, where                                        */
            Mod(x,
              Sequence(x-1)   /*       y ranges from 1 to x-1                                                            */
            ),
            0
          )
        )
      )
    )
  )
)