#include <Rdefines.h>

SEXP
R_png_dims(SEXP filename)
{
    FILE *f = fopen(CHAR(STRING_ELT(filename, 0)), "rb");
    unsigned int w, h;
    char bytes[17];
    SEXP ans;

    if(!f) {
	PROBLEM "cannot open file %s", CHAR(STRING_ELT(filename, 0))
	    ERROR;
    }
    fread(bytes, sizeof(char), 16, f);
    fscanf(f, "%u", &w);
    fscanf(f, "%u", &h);

    fprintf(stderr, "w = %u, h = %u\n", w, h);

    PROTECT(ans = NEW_NUMERIC(2));
    REAL(ans)[0] = w;
    REAL(ans)[1] = h;

    UNPROTECT(1);
    return(ans);
}
