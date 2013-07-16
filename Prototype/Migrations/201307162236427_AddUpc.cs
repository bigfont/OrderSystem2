namespace Prototype.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class AddUpc : DbMigration
    {
        public override void Up()
        {
            AddColumn("dbo.Products", "ProductUpc", c => c.String());
        }
        
        public override void Down()
        {
            DropColumn("dbo.Products", "ProductUpc");
        }
    }
}
