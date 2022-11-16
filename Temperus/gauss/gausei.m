function [x, k] = gausei(A, b, xold, eps, N)
%=====================================================================
% Solves the n by n linear system A x = b
%
% Written for comparison with the basic Gaussian elimination 
% with backward substitution algorithm
% perform row interchanges when pivot element is zero
% Optimezed for best readability, not for minimum CPU time
% Note that the eliminated elements are not explicitly set to zero
%
% INPUT:   square matrix A, right-hand-side column vector b
%
% OUTPUT:  solution vector x if the system can be solved
%=====================================================================

  tic
%---------------------------------------------------------------------
% Some checks
%---------------------------------------------------------------------
  [n,m] = size(A);
  if n~=m
    error(sprintf('Not a square matrix'));
  end
  [nb,mb] = size(b);
  if n~=nb
    error(sprintf('Number of rows in A and b not equal'));
  end
  if mb~=1
    error(sprintf('Number of columns in b not 1'));
  end

  k = 1;
%---------------------------------------------------------------------
% ITERATION LOOP
%---------------------------------------------------------------------
  while k <= N
%---------------------------------------------------------------------
%   Calculate new approximation x
%---------------------------------------------------------------------
    for i = 1:n
      y = 0;
      for j = 1:i-1
        y = y + A(i,j) * x(j);
      end
      for j = i+1:n
        y = y + A(i,j) * xold(j);
      end
      x(i) = ( b(i) - y ) / A(i,i);
    end

%---------------------------------------------------------------------
%   Check convergence with infinity norm
%---------------------------------------------------------------------
    ymax = 0;
    zmax = 0;
    for i = 1:n
      y = abs( x(i) - xold(i) );
      z = abs( x(i) );
      if ( y > ymax )
        ymax = y;
      end
      if ( z > zmax )
        zmax = z;
      end
    end

    if ( ymax < eps*zmax)
      break;
    end

%---------------------------------------------------------------------
%   Prepare for next iteration
%---------------------------------------------------------------------
    k = k+1;
    for i = 1:n
      xold(i) = x(i);
    end
  end

  if k > N
    error(sprintf('Gauss-Seidel did not converge'));
  end
  toc
