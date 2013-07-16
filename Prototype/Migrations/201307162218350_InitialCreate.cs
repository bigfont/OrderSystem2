namespace Prototype.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class InitialCreate : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.Vendors",
                c => new
                    {
                        VendorID = c.Int(nullable: false, identity: true),
                        VendorName = c.String(),
                    })
                .PrimaryKey(t => t.VendorID);
            
            CreateTable(
                "dbo.Products",
                c => new
                    {
                        ProductID = c.Int(nullable: false, identity: true),
                        ProductDescription = c.String(),
                        VendorID = c.Int(nullable: false),
                    })
                .PrimaryKey(t => t.ProductID)
                .ForeignKey("dbo.Vendors", t => t.VendorID, cascadeDelete: true)
                .Index(t => t.VendorID);
            
        }
        
        public override void Down()
        {
            DropIndex("dbo.Products", new[] { "VendorID" });
            DropForeignKey("dbo.Products", "VendorID", "dbo.Vendors");
            DropTable("dbo.Products");
            DropTable("dbo.Vendors");
        }
    }
}
