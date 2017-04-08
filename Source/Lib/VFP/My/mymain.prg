lparameters tuParameter

* If no parameters were passed, this is an install call. Otherwise, it's a
* builder call.

if pcount() = 0
	do InstallMy
else
	do MyBuilderForm
endif pcount() = 0
return
