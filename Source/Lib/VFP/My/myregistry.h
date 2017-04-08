#include MyConstants.H

* Registry constants.

#define cnSUCCESS                    0
#define cnRESERVED                   0
#define cnBUFFER_SIZE                256
	* the size of the buffer for the key value

* Registry key values.

#define cnHKEY_CLASSES_ROOT          -2147483648
#define cnHKEY_CURRENT_USER          -2147483647
#define cnHKEY_LOCAL_MACHINE         -2147483646
#define cnHKEY_USERS                 -2147483645

* Data types.

#define cnREG_SZ                     1	&& Data string
#define cnREG_EXPAND_SZ              2	&& Unicode string
#define cnREG_BINARY                 3	&& Binary
#define cnREG_DWORD                  4	&& 32-bit number
