/* tclink.c - Library code for the TCLink client API.
 *
 * TCLink Copyright (c) 2003 TrustCommerce.
 * http://www.trustcommerce.com
 * developer@trustcommerce.com
 * (626) 744-7700
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
 */

#include "Python.h"
#include "tclink.c"


#define TCLINKSEND_DOC "Send the transaction to TrustCommerce for processing.\n"

static PyObject *TCLinkSend_Py(PyObject *self, PyObject *args)
{
	PyObject *input, *output;
	PyObject *key, *value;
	int pos = 0;

	TCLinkHandle handle;
	TCLinkCon *c;
	param *p;


	if (!PyArg_ParseTuple(args, "O", &input))
		return (PyObject *)NULL;

	/* stuff the parameters */
	handle = TCLinkCreate();

	while (PyDict_Next(input, &pos, &key, &value))
		TCLinkPushParam(handle, PyString_AsString(key), PyString_AsString(value));

	Py_BEGIN_ALLOW_THREADS

	/* send the transaction */
	TCLinkSend(handle);

	Py_END_ALLOW_THREADS

	/* put the output into a dictionary */
	c = (TCLinkCon *)handle;
 	output = PyDict_New();

	for (p = c->recv_param_list; p; p = p->next)
		PyDict_SetItem(output, Py_BuildValue("s", p->name), Py_BuildValue("s", p->value));

	TCLinkDestroy(handle);


	return output;
}

#define TCLINKGETVERSION_DOC "Returns the module version string.\n"

static PyObject *TCLinkGetVersion_Py(PyObject *self, PyObject *args)
{
	char buf[64];
	return Py_BuildValue("s", TCLinkGetVersion(buf));
}


/*********************************/
/* Python Module Initialization  */
/*********************************/

static PyMethodDef trustcommerceMethods[] = {
  {"getVersion", TCLinkGetVersion_Py, METH_VARARGS, TCLINKGETVERSION_DOC},
  {"send", TCLinkSend_Py, METH_VARARGS, TCLINKSEND_DOC},
  {NULL, NULL }      /* Sentinel */
};

void inittclink(void)
{
  char *trustcommerce_documentation = "The TCLink Python module is a thin client to allow e-commerce application to run credit card transactions over the Internet.  Please visit http://www.trustcommerce.com/tclink.html for more info, or write to developer@trustcommerce.com.\n";

  TCLinkCreate((PyObject *)NULL, (PyObject *)NULL);
  Py_InitModule3("tclink", trustcommerceMethods, trustcommerce_documentation);

}

