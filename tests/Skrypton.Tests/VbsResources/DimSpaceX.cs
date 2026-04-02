_outer.Space = "+";
_outer.URLEncode = _.VAL(_outer.Space);
_outer.URLEncode = _.CONCAT(_outer.Space, "x", _outer.Space);
_outer.URLEncode = _.CONCAT("y", _outer.Space);
_outer.URLEncode = _.CONCAT(_outer.Space, "z");
_outer.URLEncode = _.VAL(_.CALLm1v1(this, _outer, "F1", _outer.Space));
_outer.URLEncode = _.VAL(_.CALLm1v1(this, _outer, "F2", _outer.Space));
_outer.URLEncode = _.VAL(_.CALLm1v1(this, _outer, "F3", _outer.Space));
_outer.URLEncode = _.VAL(_.CALLm1v1(this, _outer, "F4", _outer.Space));
_outer.URLEncode = _.VAL(_.CALLm1v1(this, _outer, "F5", _outer.Space));
public object F1()
{
    object F1_retVal = null;
    object Space = null;
    return F1_retVal;
}
public object F2(ref object Space)
{
    return _.VAL(Space);
}
public object F3(ref object Space)
{
    return _.VAL(Space);
}
public object F4(object Space)
{
    return _.VAL(Space);
}
public object F5(ref object txt)
{
    return _.CONCAT(txt, _.CALLm0argp(this, _outer.Space ?? throw new InvalidOperationException("Reference not set:Space"), _.ARGS.Val((Int16)77)));
}