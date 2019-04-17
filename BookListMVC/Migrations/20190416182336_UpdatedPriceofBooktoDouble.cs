using Microsoft.EntityFrameworkCore.Migrations;

namespace BookListMVC.Migrations
{
    public partial class UpdatedPriceofBooktoDouble : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AlterColumn<double>(
                name: "Price",
                table: "Books",
                nullable: false,
                oldClrType: typeof(string),
                oldNullable: true);
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AlterColumn<string>(
                name: "Price",
                table: "Books",
                nullable: true,
                oldClrType: typeof(double));
        }
    }
}
