function [x] = basicgecp(A, b)
%=====================================================================
% Solves the n by n linear system A x = b
%
% Performs Gaussian elimination with backward substitution
% using complete pivoting
% perform row and column interchanges to get the largest entry to
% the pivot position (ii)
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
% Initialize pointer for column interchanges
% pntr(i) = j means that column i contains the original column j
%---------------------------------------------------------------------
  for i = 1:n
    pntr(i) = i;
  end

%---------------------------------------------------------------------
% GAUSSIAN ELIMINATION
%---------------------------------------------------------------------
  for i = 1:n-1
%---------------------------------------------------------------------
%   first find maximum entry in submatrix A(i:n,i:n)
%---------------------------------------------------------------------
    maxA = 0;
    for j = i:n
      for k = i:n
        if abs( A(j,k) ) > maxA
          maxA = abs( A(j,k) );
          row = j;
          col = k;
        end;
      end;
    end;
    if abs(maxA) < 1.0e-14
      error(sprintf('Gaussian elimination impossible'));
    end

%---------------------------------------------------------------------
%   Interchange rows if necessary
%   Note: only interchange the NON-ZERO PART OF THE ROW, including b
%---------------------------------------------------------------------
    if row ~= i
      for j = i:n
        c = A(i,j);
        A(i,j) = A(row,j);
        A(row,j) = c;
      end;
      c = b(i);
      b(i) = b(row);
      b(row) = c;
    end;

%---------------------------------------------------------------------
%   Interchange columns if necessary
%   Note: interchange the WHOLE COLUMN
%         and keep track of where each column is
%---------------------------------------------------------------------
    if col ~= i
      for j = 1:n
        c = A(j,i);
        A(j,i) = A(j,col);
        A(j,col) = c;
      end;
      c = pntr(i);
      pntr(i) = pntr(col);
      pntr(col) = c;
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
      b(i) = b(i) - A(i,j) * xtmp(j);
    end;
    xtmp(i) = b(i) / A(i,i);
  end;

%---------------------------------------------------------------------
% Now reorder
%---------------------------------------------------------------------
  for i = 1:n
    x( pntr(i) ) = xtmp(i);
  end
