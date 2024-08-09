namespace TemplateMarker_with_Insert_Argument.Models
{
    public class Employee
    {
        private string m_name;
        private int m_id;
        private int m_age;

        public string Name
        {
            get
            {
                return m_name;
            }

            set
            {
                m_name = value;
            }
        }
        public int Id
        {
            get
            {
                return m_id;
            }

            set
            {
                m_id = value;
            }
        }
        public int Age
        {
            get
            {
                return m_age;
            }

            set
            {
                m_age = value;
            }
        }
    }
}
