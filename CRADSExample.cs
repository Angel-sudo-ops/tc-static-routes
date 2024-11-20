private void CreateMissingRoutes(DataGridViewRow Row, string NetID_IP, int ADSport)
{
    TwinCATAds.TwinCATCom TwinCATCom1 = new TwinCATAds.TwinCATCom();
    if (Row.Cells[0].Value is null || Row.Cells[0].Style.BackColor == Color.Green)
    { }
    else
    {
        TwinCATCom1.DisableSubScriptions = true;
        TwinCATCom1.Password = PLC_Password_TextBox.Text;// "";
        TwinCATCom1.PollRateOverride = 500;
        //TwinCATCom1.SynchronizingObject = Me;
        TwinCATCom1.TargetAMSNetID = Row.Cells[3].Value.ToString();//"10.209.80.202.1.1";
        //TwinCATCom1.TargetAMSPort = 801;
        TwinCATCom1.TargetAMSPort = Convert.ToUInt16(ADSport);
        TwinCATCom1.TargetIPAddress = Row.Cells[2].Value.ToString();//"10.209.80.202";
        TwinCATCom1.UserName = PLC_Username_TextBox.Text;// "Administrator";
        TwinCATCom1.UseStaticRoute = true;
        try
        {
            MessageBox.Show("Connection from " + NetID_IP+ " to "  + Row.Cells[2].Value.ToString() + ", responded:"
                            + TwinCATCom1.CreateRoute(Name_TextBox.Text,IPAddress.Parse(NetID_IP)));
        }
        catch (TwinCATAds.Common.PLCDriverException eex)
        {
            //MessageBox.Show("Unable to connect to IP: " + Row.Cells[2].Value.ToString()  + System.Environment.NewLine + eex.ToString());
        }
    }
}