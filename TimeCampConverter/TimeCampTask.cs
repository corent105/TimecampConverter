namespace TimecampConverter
{
    public class TimeCampTask
    {
        public string task_id { get; set; }
        public string parent_id { get; set; }
        public string assigned_by { get; set; }
        public string name { get; set; }
        public object external_task_id { get; set; }
        public object external_parent_id { get; set; }
        public string level { get; set; }
        public string archived { get; set; }
        public string tags { get; set; }
        public string budgeted { get; set; }
        public string budget_unit { get; set; }
        public string root_group_id { get; set; }
        public string billable { get; set; }
        public object note { get; set; }
        public object public_hash { get; set; }
        public string add_date { get; set; }
        public string modify_time { get; set; }
        public string color { get; set; }
        public Users users { get; set; }
        public int user_access_type { get; set; }
    }
    public class __invalid_type__1415179
    {
        public string user_id { get; set; }
        public string role_id { get; set; }
    }

    public class Users
    {
        public __invalid_type__1415179 __invalid_name__1415179 { get; set; }
    }

}