# Introduction #

There is a need to install the MySQL .NET extension in order to properly leverage this code. Any machine running this code needs this installed. For this reason, it is recommended that this be part of your installer or be run on a server setup.


# Details #

Use the following steps to install the required dependencies.
  * Download and install http://dev.mysql.com/downloads/connector/net/
  * If developing .NET apps accessing MySQL with VS2005, install the connector (adds VS2005 integration).
  * Ensure the dll is located in your bin dir at runtime and is registered as an assembly on the running maching (regasm 

&lt;filename&gt;

).