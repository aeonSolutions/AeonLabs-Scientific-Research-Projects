function [x] = basicge(A, b)
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

%---------------------------------------------------------------------
% GAUSSIAN ELIMINATION
%---------------------------------------------------------------------
  for i = 1:n-1
%---------------------------------------------------------------------
%   first find non-zero pivot element if possible
%---------------------------------------------------------------------
    p = i;
    found = 0;
    while p <= n 
      if A(p,i) ~= 0
        found = 1;
        break;
      else
        p = p+1;
      end;
    end;
    if found == 0
      error(sprintf('Gaussian elimination impossible'));
    end

%---------------------------------------------------------------------
%   Interchange rows if necessary
%---------------------------------------------------------------------
    if p ~= i
      for j = i:n
        c = A(i,j);
        A(i,j) = A(p,j);
        A(p,j) = c;
      end;
      c = b(i);
      b(i) = b(p);
      b(p) = c;
    end;

%---------------------------------------------------------------------
%   Now the elimination
%---------------------------------------------------------------------
    for j = i+1:n
      m = A(j,i) / A(i,i);
      for k = i+1:n
        A(j,k) = A(j,k) - m*A(i,k);
      end;
      b(j) = b(j) - m*b(i);
    end;
  end;

%---------------------------------------------------------------------
% BACKWARD SUBSTITUTION
% first check whether backward substitution is possible
%---------------------------------------------------------------------
  if A(n,n) == 0
    error(sprintf('Backward substitution impossible'));
  end

%---------------------------------------------------------------------
% Now the backward subsitution
%---------------------------------------------------------------------
  for i = n:-1:1
    for j = i+1:n
      b(i) = b(i) - A(i,j) * x(j);
    end;
    x(i) = b(i) / A(i,i);
  end;
