using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MinimizeWhenColseAddin
{
    public enum GWL
    {
        // Retrieves the pointer to the window procedure,
        // or a handle representing the pointer to the window procedure.
        // You must use the CallWindowProc function to call the window procedure.
        GWL_WNDPROC = (-4),

        // Retrieves a handle to the application instance.
        GWL_HINSTANCE = (-6),

        // Retrieves a handle to the parent window, if there is one.
        GWL_HWNDPARENT = (-8),

        // Retrieves the window styles.
        GWL_STYLE = (-16),

        // Retrieves the extended window styles.
        GWL_EXSTYLE = (-20),

        // Retrieves the user data associated with the window.
        // This data is intended for use by the application that created the window. Its value is initially zero.
        GWL_USERDATA = (-21),

        // Retrieves the identifier of the window.
        GWL_ID = (-12)
    }
}
