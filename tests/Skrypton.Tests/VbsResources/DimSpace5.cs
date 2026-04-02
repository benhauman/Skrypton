_outer.URLEncode = _.VAL(_.SPACE((Int16)10));
public object F5(ref object txt)
{
    object F5_retVal = null;
    object Space = null;
    Space = "__";
    F5_retVal = _.CONCAT(txt, Space);
    return F5_retVal;
}