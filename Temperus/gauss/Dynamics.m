(* Mathematica Commands for Dynamical Systems *)

BeginPackage["Dynamics`"];

Schwarzian::usage = "Schwarzian[f, x] gives the Schwarzian derivative 
of f, with respect to x.";

JacobianMatrix::usage="JacobianMatrix[{f1,...,fm},{x1,...xn}] 
calculates the Jacobian matrix.";

Cobweb::usage = "Cobweb[f, x0, n, k] produces a cobweb plot of 
iterations of a one-dimensional map named f, initialized at x0. 
Terms n through n+k are plotted. The plot range is determined 
automatically, unless specified explicitly by an optional fifth 
argument: Cobweb[f, x0, n, k, {xmin,xmax}]";

Trajectory::usage = "Trajectory[f, x0, n, k] produces a plot of the 
trajectory (orbit) sequence of the iterations of a two-dimensional 
map named f, initialized at x0.  Terms n through n+k are plotted.  f 
and x0 must be two dimensional (lists of length 2).  The size of 
plotted points can be controlled by an optimal fifth argument: 
Trajectory[f, x0, n, k, ptsize].  The default is  ptsize=008.";

Begin["`Private`"];

Schwarzian[f_,x_]:=(D[f,{x,3}]/D[f,x])-(3/2)*(D[f,{x,2}]/D[f,x])^2

JacobianMatrix[fcns_List,vars_List]:=Outer[D,fcns,vars]

Trajectory[f_, x0_, initial_, orbitlength_, ptsize_:.008] :=
Module[{start},start=Nest[f,N[x0],initial];
  Show[Graphics[{PointSize[ptsize],Map[Point,
  NestList[f,start,orbitlength]]}],
  AspectRatio->1,PlotRange->All,Frame->True]];

Cobweb[f_, x0_, initial_, orbitlength_, xrange_:0] :=
Module[{start, orbit, fxplt, lines, xmin, xmax, border},
  start = Nest[f, N[x0], initial];
  orbit = NestList[f, start, orbitlength];
  xmin=Min[orbit];
  xmax=Max[orbit];
  border=.2*(xmax-xmin);
  If[Length[xrange]==0,{xmin,xmax}={Min[0,xmin-border],Max[1,xmax+border]},{xmin,xmax}=xrange];
  fxplt = Plot[f[x], {x, xmin,xmax}, DisplayFunction->Identity];
  lines = Line[Rest[Partition[Flatten[Transpose[{orbit,orbit}]], 2, 1]]];

  Show[fxplt,Graphics[{{Thickness[.0001], 
PointSize[.02],lines,Point[{start, f[start]}],
       Line[{{xmin, xmin}, {xmax, xmax}}]}}], AxesOrigin->{xmin, xmin},
       DisplayFunction->$DisplayFunction, PlotRange->{{xmin, 
xmax},{xmin, xmax}}]
];

End[ ];
EndPackage[ ];
Null