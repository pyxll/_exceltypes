#include "exceltypes.h"
#include "PyIRTDUpdateEvent.h"

/* List of module functions */
static struct PyMethodDef exceltypes_methods[] =
{
	{ NULL, NULL },
};

static const PyCom_InterfaceSupportInfo g_interfaceSupportData[] =
{
    PYCOM_INTERFACE_CLIENT_ONLY(RTDUpdateEvent)
};

/* Module initialisation */
PYWIN_MODULE_INIT_FUNC(_exceltypes)
{
    PyWinGlobals_Ensure();
	PYWIN_MODULE_INIT_PREPARE(_exceltypes, exceltypes_methods, "A module wrapping Excel COM interfaces");
    PyCom_RegisterExtensionSupport(dict, g_interfaceSupportData, sizeof(g_interfaceSupportData)/sizeof(PyCom_InterfaceSupportInfo));
    PYWIN_MODULE_INIT_RETURN_SUCCESS;
}
