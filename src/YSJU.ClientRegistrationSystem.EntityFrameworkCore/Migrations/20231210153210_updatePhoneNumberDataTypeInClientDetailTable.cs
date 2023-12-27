using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace YSJU.ClientRegistrationSystem.Migrations
{
    /// <inheritdoc />
    public partial class updatePhoneNumberDataTypeInClientDetailTable : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AlterColumn<string>(
                name: "PhoneNumber",
                table: "ClientDetails",
                type: "character varying(16)",
                maxLength: 16,
                nullable: false,
                oldClrType: typeof(int),
                oldType: "integer",
                oldMaxLength: 16);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AlterColumn<int>(
                name: "PhoneNumber",
                table: "ClientDetails",
                type: "integer",
                maxLength: 16,
                nullable: false,
                oldClrType: typeof(string),
                oldType: "character varying(16)",
                oldMaxLength: 16);
        }
    }
}
