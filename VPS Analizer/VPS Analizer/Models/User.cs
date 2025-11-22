namespace VPS_Analizer.Models
{
    public class UserInfo
    {
        public Guid UserId { get; set; }
        public string VpsId { get; set; }
        public string VpsIp { get; set; }
        public string VpsPassword { get; set; }
        public DateTime LastCheckedTime { get; set; }
        public string ServerRam { get; set; }
        public string ServerCpu { get; set; }
        public string ClientLogin { get; set; }
        public string AccountBalance { get; set; }
        public string AccountEquity { get; set; }
        public bool RobotStatus { get; set; }
        public string ProblemDescription { get; set; }
    }
}