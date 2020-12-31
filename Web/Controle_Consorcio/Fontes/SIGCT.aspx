<%@ Page Language="C#" AutoEventWireup="true" CodeFile="SIGCT.aspx.cs" Inherits="SIGCT" %>
<html>
    <head>
        <meta charset="utf-8">
        <meta http-equiv="X-UA-Compatible" content="IE=edge">
        <title>Cadastro de Funcionários</title>
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
        <style type="text/css">
            .auto-style1 {
                width: 5px;
                height: 43px;
            }
            .auto-style2 {
                height: 43px;
                text-align: right;
            }
            .auto-style4 {
                height: 30px;
                width: 303px;
                text-align: right;
            }
            .auto-style5 {
                width: 303px;
                height: 43px;
                text-align: right;
            }
            .auto-style6 {
                width: 60px;
                height: 43px;
                text-align: right;
            }
            .auto-style7 {
                width: 50px;
                height: 43px;
            }
            .auto-style8 {
                width: 99px;
                height: 43px;
            }
            .auto-style9 {
                height: 30px;
                width: 99px;
            }
        </style>
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
                </aside>
                <!-------------------------------------------------------------------------------------->
                <div class="content-wrapper">
                    <section class="content-header">
                        <h2><Center> Controle de Demandas - SIGCT </Center> </h2>
                    </section>
                    <section class="content">
                        <div class="row">
                            <table>
                                <tr style="width:10px;height:10px">
                                    <td style="width:5px;height:5px"></td>
                                    <td style="width:25px;height:20px" class="text-right">
                                        <label> Líder </label>
                                    </td>
                                    <td style="width:5px;height:5px">
                                    <td style="width:170px;height:60px">
                                        <select id="CboLider" class="form-control" runat="server" name="D1"></select>
                                    </td>
                                    <td style="width:5px;height:5px">
                                    <td style="width:5px;height:5px"></td>
                                    <td style="width:5px;height:5px"></td>
                                    <td style="width:25px;height:20px" class="text-right">
                                        <label> Período </label>
                                    </td>
                                    <td style="width:170px;height:60px">
                                        <select id="CboPeriodo" class="form-control" runat="server" name="D1"></select>
                                    </td>
                                    <td style="width:1px;height:1px"></td>
                                    <td style="width:50px;height:30px">
                                        <asp:Button ID="CmdPesquisar" runat="server" Text="Pesquisar" class="btn btn-block btn-primary" Width="100px"/>
                                    <td style="width:220px;height:30px"></td>
                                </tr>
                                <tr style="width:10px;height:10px">
                                    <td style="width:5px;height:5px"></td>
                                    <td style="width:5px;height:5px" class="text-right">
                                        <label> Contrato </label>
                                    </td>
                                    <td style="width:5px;height:5px"></td>
                                    <td style="width:160px;height:60px">
                                        <select id="CboContrato" class="form-control" runat="server" name="D1"></select>
                                    </td>
                                    <td style="width:5px;height:5px"></td>
                                    <td style="width:150px;height:20px" class="text-right">
                                        <label> Demanda </label>
                                    </td>
                                    <td style="width:5px;height:5px">
                                    <td style="width:150px;height:30px" class="text-right">
                                        <asp:TextBox ID="TxtDemanda" runat="server" Width="180px" Font-Size="Small"></asp:TextBox>
                                    </td>
                                    <td style="width:140px;height:20px" class="text-right">
                                        <label> Ordem de Serviço </label>
                                    </td>
                                    <td style="width:5px;height:5px">
                                    <td style="width:100px;height:30px" class="text-right">
                                        <asp:TextBox ID="TxtOrdemServico" runat="server" Width="100px" Font-Size="Small"></asp:TextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td style="width:5px;height:5px"></td>
                                    <td style="width:100px;height:20px" class="text-right">
                                        <label> Qtde UST's </label>
                                    </td>
                                    <td style="width:5px;height:5px">
                                    <td style="width:100px;height:30px" class="text-right">
                                        <asp:TextBox ID="TxtQtdeUST" runat="server" Width="100px" Font-Size="Small"></asp:TextBox>
                                    </td>
                                    <td style="width:10px;height:20px" class="text-right"></td>
                                    <td style="width:10px;height:5px" class="text-right">
                                        <label> Valor </label>
                                    </td>
                                    <td style="width:5px;height:5px">
                                    <td style="width:100px;height:40px">
                                        <asp:TextBox ID="TextBox1" runat="server" Width="140px" Font-Size="Small"></asp:TextBox>
                                    </td>
                                    <td style="width:150px;height:20px" class="text-right">
                                        <label> Data de Entrega </label>
                                    </td>
                                    <td style="width:5px;height:5px">
                                    <td style="width:150px;height:30px" class="text-right">
                                        <asp:TextBox ID="TxtDataEntrega" runat="server" Width="180px" Font-Size="Small"></asp:TextBox>
                                    </td>
                                    <td style="width:130px;height:20px" class="text-right">
                                        <label> Projeto </label>
                                    </td>
                                    <td style="width:5px;height:5px">
                                    <td style="width:100px;height:30px" class="text-right">
                                        <select id="CboProjeto" class="form-control" runat="server" name="D1"></select>
                                    </td>
                                </tr>
                            </table>
                            <table>
                                <tr>
                                    <td style="width:5px;height:20px" class="text-right"></td>
                                    <td style="width:70px;height:20px" class="text-right">
                                        <label> Item do Contrato </label>
                                    </td>
                                    <td style="width:5px;height:5px">
                                    <td style="width:200px;height:60px">
                                        <select id="CboItemContrato" class="form-control" runat="server" name="D1"></select>
                                    </td>
                                    <td style="width:5px;height:20px" class="text-right"></td>
                                    <td style="width:70px;height:20px" class="text-right">
                                        <label> OBS </label>
                                    </td>
                                    <td style="width:5px;height:10px" class="text-right"></td>
                                    <td style="width:100px;height:30px" class="text-right">
                                        <asp:TextBox ID="TxtOBS" runat="server" Width="730px" Font-Size="Small" Height="50px" TextMode="MultiLine"></asp:TextBox>
                                    </td>
                                </tr>
                            </table>
                            <table>
                                <tr>
                                    <asp:Label ID="LblExportacao" runat="server" Text=""></asp:Label>
                                </tr>
                                <tr>
                                    <td class="auto-style9"></td>
                                    <td style="width:50px;height:30px">
                                        <asp:Button ID="CmdNovo" runat="server" Text="Novo" class="btn btn-block btn-primary" Width="100px"/>
                                    <td style="width:220px;height:30px"></td>
                                    <td style="width:220px;height:30px"></td>
                                    <td style="width:50px;height:30px">
                                        <asp:Button ID="CmdSalvar" runat="server" Text="Salvar" class="btn btn-block btn-primary" Width="100px"/>
                                    <td style="width:220px;height:30px"></td>
                                    <td style="width:230px;height:30px"></td>
                                    <td style="width:50px;height:30px">
                                        <asp:Button ID="CmdLimpar" runat="server" Text="Limpar Seleção" Width="140px" class="btn btn-block btn-primary"/>
                                    </td>
                                </tr>
                            </table>
                            <table>
                                <tr>
    	                            <asp:GridView ID="Grid_Funcionarios_Novo" runat="server" AutoGenerateColumns="false" class="table table-bordered table-striped">
	    	                            <Columns>
			                                <asp:TemplateField>
				                                <ItemTemplate>
					                                <asp:Button ID="btnEditar" class="btn btn-block btn-primary btn-xs" runat="server" CommandName="Editar" Text="Editar"
					                                CommandArgument='<%# DataBinder.Eval(Container.DataItem, "Matricula")%>' />
				                                </ItemTemplate>
			                                </asp:TemplateField>
		    	                            <asp:BoundField DataField="Matricula" HeaderText="Matricula" />
			                                <asp:BoundField DataField="Nome" HeaderText="Nome" />
                                            <asp:BoundField DataField="Situacao" HeaderText="Situacao" />
			                                <asp:BoundField DataField="Horario_Escala_Trabalho" HeaderText="Horário de Trabalho" />
                                            <asp:BoundField DataField="Funcao" HeaderText="Função" />
                                            <asp:BoundField DataField="Data_Admissao" HeaderText="Data de Admissão" />
		                                </Columns>
	                                </asp:GridView>
                                </tr>
                            </table>
                        </div>
                    </section>
                </div>
                <footer class="main-footer">
                    <div class="pull-right hidden-xs">
                        <b>Versionen-xs">
                            <div class="pull-right hidden-xs">
                        <b>Versionen-xs">
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
                &nbsp;<p>
                Some information about this general settings option
                </p>
                </div>

                <div class="form-group">
                <label class="control-sidebar-subheading">
                Allow mail redirect
                <input type="checkbox" class="pull-right" checked>
                </label>
                &nbsp;<p>
                Other sets of options are available
                </p>
                </div>

                <div class="form-group">
                <label class="control-sidebar-subheading">
                Expose author name in posts
                <input type="checkbox" class="pull-right" checked>
                </label>
                &nbsp;<p>
                Allow the user to show his name in blog posts
                </p>
                </div>

                <h3 class="control-sidebar-heading">Chat Settings</h3>

                <div class="form-group">
                <label class="control-sidebar-subheading">
                Show me as online
                <input type="checkbox" class="pull-right" checked>
                </label>
                &nbsp;</div>

                <div class="form-group">
                <label class="control-sidebar-subheading">
                Turn off notifications
                <input type="checkbox" class="pull-right">
                </label>
                &nbsp;</div>

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
        </form>
    </body>
</html>