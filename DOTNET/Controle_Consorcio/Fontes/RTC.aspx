<%@ Page Language="C#" AutoEventWireup="true" CodeFile="RTC.aspx.cs" Inherits="RTC" %>
<!DOCTYPE html>
<html>
    <head>
        <meta charset="utf-8">
        <meta http-equiv="X-UA-Compatible" content="IE=edge">
        <title>Controle Consórcio</title>
        <meta content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no" name="viewport">
        <link rel="stylesheet" href="../../bootstrap/css/bootstrap.min.css">
        <link rel="stylesheet" href="../../plugins/fonts/css/font-awesome.min.css">
        <link rel="stylesheet" href="../../plugins/daterangepicker/daterangepicker-bs3.css">
        <link rel="stylesheet" href="../../plugins/iCheck/all.css">
        <link rel="stylesheet" href="../../plugins/colorpicker/bootstrap-colorpicker.min.css">
        <link rel="stylesheet" href="../../plugins/timepicker/bootstrap-timepicker.min.css">
        <link rel="stylesheet" href="../../plugins/select2/select2.min.css">
        <link rel="stylesheet" href="../../dist/css/AdminLTE.min.css">
        <link rel="stylesheet" href="../../dist/css/skins/_all-skins.min.css">
    </head>
    <body class="hold-transition skin-blue sidebar-mini">
        <form id="form1" runat="server">
            <div class="wrapper">
                <header class="main-header">
                    <a href="#" class="logo">
                    <span class="logo-mini"><b>A</b>LT</span>
                    <span class="logo-lg"><b>Gestão</b> Consórcio</span>
                    </a>
                    <nav class="navbar navbar-static-top" role="navigation">
                    <a href="#" class="sidebar-toggle" data-toggle="offcanvas" role="button">
                    <span class="sr-only">Toggle navigation</span>
                    </a>
                    </nav>
                </header>
                <aside class="main-sidebar">
                    <section class="sidebar">
                        <!-- Menu Lateral -->
                        <ul class="sidebar-menu">
                            <li>
                                <a href="Default.aspx">
                                    <i class="fa fa-th"></i> <span>Painel Principal</span>
                                </a>
                            </li>
                            <li class="treeview">
                                <a href="#">
                                    <i class="fa fa-pie-chart"></i>
                                    <span>Cadastros</span>
                                    <i class="fa fa-angle-left pull-right"></i>
                                </a>
                                <ul class="treeview-menu">
                                    <li><a href="Cadastro_Funcionarios.aspx"><i class="fa fa-circle-o"></i> Funcionários </a></li>
                                </ul>
                            </li>
                            <li class="treeview">
                                <a href="#">
                                    <i class="fa fa-pie-chart"></i>
                                    <span>Gerencial</span>
                                    <i class="fa fa-angle-left pull-right"></i>
                                </a>
                                <ul class="treeview-menu">
                                    <li><a href="Consumo_USTs.aspx"><i class="fa fa-circle-o"></i>Consumo UST's</a></li>
                                    <li><a href="RTC.aspx"><i class="fa fa-circle-o"></i>Demandas RTC</a></li>
                                    <li><a href="SIGCT.aspx"><i class="fa fa-circle-o"></i>Demandas Grupo 1 (SIGCT)</a></li>
                                </ul>
                            </li>
                            <li class="treeview">
                                <a href="#">
                                    <i class="fa fa-laptop"></i>
                                    <span>Manutenção</span>
                                    <i class="fa fa-angle-left pull-right"></i>
                                </a>
                                <ul class="treeview-menu">
                                    <li><a href="Manutencao.aspx"><i class="fa fa-circle-o"></i>Demandas / Serviços </a></li>
                                </ul>
                            </li>
                        </ul>
                    </section>
                    </aside>
            <!-------------------------------------------------------------------------------------->
                </aside>
                <div class="content-wrapper">
                    <section class="content-header">
                        <h1> Análise Demandas - RTC</h1>
                        <ol class="breadcrumb">
                            <li><a href="#"><i class="fa fa-dashboard"></i> Home</a></li>
                            <li><a href="#">Forms</a></li>
                            <li class="active">Análise Demandas - RTC</li>
                        </ol>
                    </section>
                    <!-- Filtros -->
                    <section class="content">
                        <div class="row">
                            <table>
                                <tr style="width:1000px;height:100px">
                                    <td style="width:10px;height=10px"></td>
                                    <td style="width:300px;height:30px">
                                        <label> Selecione a Unidade </label>
                                        <select multiple class="form-control" ID="Opt_Unidade" runat="server">
                                            <option > CEDES / SP </option>
                                            <option > CEDES / BR </option>
                                            <option > CEDES / RJ </option>
                                        </select>
                                    </td>
                                    <td style="width:5px;height=10px"></td>
                                    <td style="width:120px;height:20px">
                                        <label>Contrato
                                        <select id="CboContrato" class="form-control" runat="server" name="D1">
                                            <option></option>
                                            <option>CTMARG</option>
                                            <option>CTMONSI</option>
                                        </select></label><i class="fa fa-calendar"></i>
                                    </td>
                                    <td style="width:5px;height=10px"></td>
                                    <td style="width:120px;height:20px">
                                        <label>Período Previsto
                                        <select id="CboPeriodoPrevisto" class="form-control" runat="server" name="D1">
                                            <option></option></select></label><i class="fa fa-calendar"></i>
                                    </td>
                                    <td style="width:5px;height=10px"></td>
                                    <td style="width:120px;height:20px">
                                        <label>Período Real 
                                        <select id="CboPeriodoReal" class="form-control" runat="server" name="D1">
                                            <option></option></select></label><i class="fa fa-calendar"></i>
                                    </td>
                                    <td style="width:5px;height=10px"></td>
                                    <td style="width:300px;height=300px">
                                        <label>Período (Data de Criação)</label>
                                        <i class="fa fa-calendar"></i>
                                        <input type="text" ID="TxtPeriodo" class="form-control pull-right" runat="server">
                                    </td>
                                    <td style="width:5px;height=10px"></td>
                                    <td style="width:300px; height:30px">
                                        <label>Tipo de Consulta</label>
                                        <select id="CboTipoConsulta" class="form-control" runat="server" OnChange="CboTipoConsulta_Change">
                                            <option>Demandas sem tag</option>
                                            <option>Demandas atrasadas</option>
                                            <option>Período Real x Período Previsto</option>
                                        </select>
                                    </td>
                                    <td style="width:300px; height:10px">
                                        <asp:Button ID="CmdPesquisar" runat="server" Text="Pesquisar..." class="btn btn-app" OnClick="CmdPesquisar_Click"/>
                                    </td>
                                </tr>
                                <tr>
                                    <td style="width:5px;height=10px"></td>
                                    <td>
                                        <asp:Label ID="LblResumoConsulta" runat="server" Text="Label"></asp:Label>
                                    </td>
                                </tr>
                                <tr>
                                    <asp:GridView ID="Grid_Demandas" runat="server" class="table table-bordered table-striped" AllowSorting="true" OnSorting="OrdenaGridDemandas"></asp:GridView>
                                </tr>
                            </table>
                        </div>
                    </section>
                </div>
                <footer class="main-footer">
                    <div class="pull-right hidden-xs">
                        <b>Version</b> 1.0.0
                    </div>
                    <strong>Copyright &copy; 2014-2015 <a href="http://almsaeedstudio.com">Almsaeed Studio</a>.</strong> All rights reserved.
                </footer>
                <aside class="control-sidebar control-sidebar-dark">
                <!-- Create the tabs -->
                <ul class="nav nav-tabs nav-justified control-sidebar-tabs">
                <li><a href="#control-sidebar-home-tab" data-toggle="tab"><i class="fa fa-home"></i></a></li>
                <li><a href="#control-sidebar-settings-tab" data-toggle="tab"><i class="fa fa-gears"></i></a></li>
                </ul>
                <!-- Tab panes -->
                <div class="tab-content">
                <!-- Home tab content -->
                <div class="tab-pane" id="control-sidebar-home-tab">
                <h3 class="control-sidebar-heading">Recent Activity</h3>
                <ul class="control-sidebar-menu">
                <li>
                <a href="javascript::;">
                <i class="menu-icon fa fa-birthday-cake bg-red"></i>
                <div class="menu-info">
                <h4 class="control-sidebar-subheading">Langdon's Birthday</h4>
                <p>Will be 23 on April 24th</p>
                </div>
                </a>
                </li>
                <li>
                <a href="javascript::;">
                <i class="menu-icon fa fa-user bg-yellow"></i>
                <div class="menu-info">
                <h4 class="control-sidebar-subheading">Frodo Updated His Profile</h4>
                <p>New phone +1(800)555-1234</p>
                </div>
                </a>
                </li>
                <li>
                <a href="javascript::;">
                <i class="menu-icon fa fa-envelope-o bg-light-blue"></i>
                <div class="menu-info">
                <h4 class="control-sidebar-subheading">Nora Joined Mailing List</h4>
                <p>nora@example.com</p>
                </div>
                </a>
                </li>
                <li>
                <a href="javascript::;">
                <i class="menu-icon fa fa-file-code-o bg-green"></i>
                <div class="menu-info">
                <h4 class="control-sidebar-subheading">Cron Job 254 Executed</h4>
                <p>Execution time 5 seconds</p>
                </div>
                </a>
                </li>
                </ul><!-- /.control-sidebar-menu -->

                <h3 class="control-sidebar-heading">Tasks Progress</h3>
                <ul class="control-sidebar-menu">
                <li>
                <a href="javascript::;">
                <h4 class="control-sidebar-subheading">
                Custom Template Design
                <span class="label label-danger pull-right">70%</span>
                </h4>
                <div class="progress progress-xxs">
                <div class="progress-bar progress-bar-danger" style="width: 70%"></div>
                </div>
                </a>
                </li>
                <li>
                <a href="javascript::;">
                <h4 class="control-sidebar-subheading">
                Update Resume
                <span class="label label-success pull-right">95%</span>
                </h4>
                <div class="progress progress-xxs">
                <div class="progress-bar progress-bar-success" style="width: 95%"></div>
                </div>
                </a>
                </li>
                <li>
                <a href="javascript::;">
                <h4 class="control-sidebar-subheading">
                Laravel Integration
                <span class="label label-warning pull-right">50%</span>
                </h4>
                <div class="progress progress-xxs">
                <div class="progress-bar progress-bar-warning" style="width: 50%"></div>
                </div>
                </a>
                </li>
                <li>
                <a href="javascript::;">
                <h4 class="control-sidebar-subheading">
                Back End Framework
                <span class="label label-primary pull-right">68%</span>
                </h4>
                <div class="progress progress-xxs">
                <div class="progress-bar progress-bar-primary" style="width: 68%"></div>
                </div>
                </a>
                </li>
                </ul>
                </div>
                <div class="tab-pane" id="control-sidebar-stats-tab">Stats Tab Content</div><!-- /.tab-pane -->
                <div class="tab-pane" id="control-sidebar-settings-tab">
                <h3 class="control-sidebar-heading">General Settings</h3>
                <div class="form-group">
                <label class="control-sidebar-subheading">
                Report panel usage
                <input type="checkbox" class="pull-right" checked>
                </label>
                <p>
                Some information about this general settings option
                </p>
                </div>

                <div class="form-group">
                <label class="control-sidebar-subheading">
                Allow mail redirect
                <input type="checkbox" class="pull-right" checked>
                </label>
                <p>
                Other sets of options are available
                </p>
                </div>

                <div class="form-group">
                <label class="control-sidebar-subheading">
                Expose author name in posts
                <input type="checkbox" class="pull-right" checked>
                </label>
                <p>
                Allow the user to show his name in blog posts
                </p>
                </div>

                <h3 class="control-sidebar-heading">Chat Settings</h3>

                <div class="form-group">
                <label class="control-sidebar-subheading">
                Show me as online
                <input type="checkbox" class="pull-right" checked>
                </label>
                </div>

                <div class="form-group">
                <label class="control-sidebar-subheading">
                Turn off notifications
                <input type="checkbox" class="pull-right">
                </label>
                </div>

                <div class="form-group">
                <label class="control-sidebar-subheading">
                Delete chat history
                <a href="javascript::;" class="text-red pull-right"><i class="fa fa-trash-o"></i></a>
                </label>
                </div>
                </div>
                </div>
                </aside>
                <div class="control-sidebar-bg"></div>
                </div>

                <script src="../../plugins/jQuery/jQuery-2.1.4.min.js"></script>
                <script src="../../bootstrap/js/bootstrap.min.js"></script>
                <script src="../../plugins/select2/select2.full.min.js"></script>
                <script src="../../plugins/input-mask/jquery.inputmask.js"></script>
                <script src="../../plugins/input-mask/jquery.inputmask.date.extensions.js"></script>
                <script src="../../plugins/input-mask/jquery.inputmask.extensions.js"></script>
                <script src="../../plugins/moment/moment.min.js"></script>
                <script src="../../plugins/daterangepicker/daterangepicker.js"></script>
                <script src="../../plugins/colorpicker/bootstrap-colorpicker.min.js"></script>
                <script src="../../plugins/timepicker/bootstrap-timepicker.min.js"></script>
                <script src="../../plugins/slimScroll/jquery.slimscroll.min.js"></script>
                <script src="../../plugins/iCheck/icheck.min.js"></script>
                <script src="../../plugins/fastclick/fastclick.min.js"></script>
                <script src="../../dist/js/app.min.js"></script>
                <script src="../../dist/js/demo.js"></script>
                <script>
                $(function () {
                    //Initialize Select2 Elements
                    $(".select2").select2();

                    //Datemask dd/mm/yyyy
                    $("#datemask").inputmask("dd/mm/yyyy", { "placeholder": "dd/mm/yyyy" });
                    //Datemask2 mm/dd/yyyy
                    $("#datemask2").inputmask("dd/mm/yyyy", { "placeholder": "dd/mm/yyyy" });
                    //Money Euro
                    $("[data-mask]").inputmask();

                    //Date range picker
                    $('#TxtPeriodo').daterangepicker();
                    //Date range as a button
                    $('#daterange-btn').daterangepicker(
                    {
                    ranges: {
                    'Today': [moment(), moment()],
                    'Yesterday': [moment().subtract(1, 'days'), moment().subtract(1, 'days')],
                    'Last 7 Days': [moment().subtract(6, 'days'), moment()],
                    'Last 30 Days': [moment().subtract(29, 'days'), moment()],
                    'This Month': [moment().startOf('month'), moment().endOf('month')],
                    'Last Month': [moment().subtract(1, 'month').startOf('month'), moment().subtract(1, 'month').endOf('month')]
                    },
                    startDate: moment().subtract(29, 'days'),
                    endDate: moment()
                    },
                function (start, end) {
                    $('#reportrange span').html(start.format('DD MM, YYYY') + ' - ' + end.format('DD MM, YYYY'));
                    }
                    );

                    //iCheck for checkbox and radio inputs
                    $('input[type="checkbox"].minimal, input[type="radio"].minimal').iCheck({
                    checkboxClass: 'icheckbox_minimal-blue',
                    radioClass: 'iradio_minimal-blue'
                    });
                    //Red color scheme for iCheck
                    $('input[type="checkbox"].minimal-red, input[type="radio"].minimal-red').iCheck({
                    checkboxClass: 'icheckbox_minimal-red',
                    radioClass: 'iradio_minimal-red'
                    });
                    //Flat red color scheme for iCheck
                    $('input[type="checkbox"].flat-red, input[type="radio"].flat-red').iCheck({
                    checkboxClass: 'icheckbox_flat-green',
                    radioClass: 'iradio_flat-green'
                    });

                    //Colorpicker
                    $(".my-colorpicker1").colorpicker();
                    //color picker with addon
                    $(".my-colorpicker2").colorpicker();

                    //Timepicker
                    $(".timepicker").timepicker({
                    showInputs: false
                    });
                });
            </script>
        </form>
    </body>
</html>