using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace YSJU.ClientRegistrationSystem.Migrations
{
    /// <inheritdoc />
    public partial class InititalDBSetUp : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropForeignKey(
                name: "FK_AbpAuditLogActions_AbpAuditLogs_AuditLogId",
                table: "AbpAuditLogActions");

            migrationBuilder.DropForeignKey(
                name: "FK_AbpEntityChanges_AbpAuditLogs_AuditLogId",
                table: "AbpEntityChanges");

            migrationBuilder.DropForeignKey(
                name: "FK_AbpEntityPropertyChanges_AbpEntityChanges_EntityChangeId",
                table: "AbpEntityPropertyChanges");

            migrationBuilder.DropForeignKey(
                name: "FK_AbpOrganizationUnitRoles_AbpOrganizationUnits_OrganizationU~",
                table: "AbpOrganizationUnitRoles");

            migrationBuilder.DropForeignKey(
                name: "FK_AbpOrganizationUnitRoles_AbpRoles_RoleId",
                table: "AbpOrganizationUnitRoles");

            migrationBuilder.DropForeignKey(
                name: "FK_AbpRoleClaims_AbpRoles_RoleId",
                table: "AbpRoleClaims");

            migrationBuilder.DropForeignKey(
                name: "FK_AbpTenantConnectionStrings_AbpTenants_TenantId",
                table: "AbpTenantConnectionStrings");

            migrationBuilder.DropForeignKey(
                name: "FK_AbpUserClaims_AbpUsers_UserId",
                table: "AbpUserClaims");

            migrationBuilder.DropForeignKey(
                name: "FK_AbpUserLogins_AbpUsers_UserId",
                table: "AbpUserLogins");

            migrationBuilder.DropForeignKey(
                name: "FK_AbpUserOrganizationUnits_AbpOrganizationUnits_OrganizationU~",
                table: "AbpUserOrganizationUnits");

            migrationBuilder.DropForeignKey(
                name: "FK_AbpUserOrganizationUnits_AbpUsers_UserId",
                table: "AbpUserOrganizationUnits");

            migrationBuilder.DropForeignKey(
                name: "FK_AbpUserRoles_AbpRoles_RoleId",
                table: "AbpUserRoles");

            migrationBuilder.DropForeignKey(
                name: "FK_AbpUserRoles_AbpUsers_UserId",
                table: "AbpUserRoles");

            migrationBuilder.DropForeignKey(
                name: "FK_AbpUserTokens_AbpUsers_UserId",
                table: "AbpUserTokens");

            migrationBuilder.AddForeignKey(
                name: "FK_AbpAuditLogActions_AbpAuditLogs_AuditLogId",
                table: "AbpAuditLogActions",
                column: "AuditLogId",
                principalTable: "AbpAuditLogs",
                principalColumn: "Id",
                onDelete: ReferentialAction.Restrict);

            migrationBuilder.AddForeignKey(
                name: "FK_AbpEntityChanges_AbpAuditLogs_AuditLogId",
                table: "AbpEntityChanges",
                column: "AuditLogId",
                principalTable: "AbpAuditLogs",
                principalColumn: "Id",
                onDelete: ReferentialAction.Restrict);

            migrationBuilder.AddForeignKey(
                name: "FK_AbpEntityPropertyChanges_AbpEntityChanges_EntityChangeId",
                table: "AbpEntityPropertyChanges",
                column: "EntityChangeId",
                principalTable: "AbpEntityChanges",
                principalColumn: "Id",
                onDelete: ReferentialAction.Restrict);

            migrationBuilder.AddForeignKey(
                name: "FK_AbpOrganizationUnitRoles_AbpOrganizationUnits_OrganizationU~",
                table: "AbpOrganizationUnitRoles",
                column: "OrganizationUnitId",
                principalTable: "AbpOrganizationUnits",
                principalColumn: "Id",
                onDelete: ReferentialAction.Restrict);

            migrationBuilder.AddForeignKey(
                name: "FK_AbpOrganizationUnitRoles_AbpRoles_RoleId",
                table: "AbpOrganizationUnitRoles",
                column: "RoleId",
                principalTable: "AbpRoles",
                principalColumn: "Id",
                onDelete: ReferentialAction.Restrict);

            migrationBuilder.AddForeignKey(
                name: "FK_AbpRoleClaims_AbpRoles_RoleId",
                table: "AbpRoleClaims",
                column: "RoleId",
                principalTable: "AbpRoles",
                principalColumn: "Id",
                onDelete: ReferentialAction.Restrict);

            migrationBuilder.AddForeignKey(
                name: "FK_AbpTenantConnectionStrings_AbpTenants_TenantId",
                table: "AbpTenantConnectionStrings",
                column: "TenantId",
                principalTable: "AbpTenants",
                principalColumn: "Id",
                onDelete: ReferentialAction.Restrict);

            migrationBuilder.AddForeignKey(
                name: "FK_AbpUserClaims_AbpUsers_UserId",
                table: "AbpUserClaims",
                column: "UserId",
                principalTable: "AbpUsers",
                principalColumn: "Id",
                onDelete: ReferentialAction.Restrict);

            migrationBuilder.AddForeignKey(
                name: "FK_AbpUserLogins_AbpUsers_UserId",
                table: "AbpUserLogins",
                column: "UserId",
                principalTable: "AbpUsers",
                principalColumn: "Id",
                onDelete: ReferentialAction.Restrict);

            migrationBuilder.AddForeignKey(
                name: "FK_AbpUserOrganizationUnits_AbpOrganizationUnits_OrganizationU~",
                table: "AbpUserOrganizationUnits",
                column: "OrganizationUnitId",
                principalTable: "AbpOrganizationUnits",
                principalColumn: "Id",
                onDelete: ReferentialAction.Restrict);

            migrationBuilder.AddForeignKey(
                name: "FK_AbpUserOrganizationUnits_AbpUsers_UserId",
                table: "AbpUserOrganizationUnits",
                column: "UserId",
                principalTable: "AbpUsers",
                principalColumn: "Id",
                onDelete: ReferentialAction.Restrict);

            migrationBuilder.AddForeignKey(
                name: "FK_AbpUserRoles_AbpRoles_RoleId",
                table: "AbpUserRoles",
                column: "RoleId",
                principalTable: "AbpRoles",
                principalColumn: "Id",
                onDelete: ReferentialAction.Restrict);

            migrationBuilder.AddForeignKey(
                name: "FK_AbpUserRoles_AbpUsers_UserId",
                table: "AbpUserRoles",
                column: "UserId",
                principalTable: "AbpUsers",
                principalColumn: "Id",
                onDelete: ReferentialAction.Restrict);

            migrationBuilder.AddForeignKey(
                name: "FK_AbpUserTokens_AbpUsers_UserId",
                table: "AbpUserTokens",
                column: "UserId",
                principalTable: "AbpUsers",
                principalColumn: "Id",
                onDelete: ReferentialAction.Restrict);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropForeignKey(
                name: "FK_AbpAuditLogActions_AbpAuditLogs_AuditLogId",
                table: "AbpAuditLogActions");

            migrationBuilder.DropForeignKey(
                name: "FK_AbpEntityChanges_AbpAuditLogs_AuditLogId",
                table: "AbpEntityChanges");

            migrationBuilder.DropForeignKey(
                name: "FK_AbpEntityPropertyChanges_AbpEntityChanges_EntityChangeId",
                table: "AbpEntityPropertyChanges");

            migrationBuilder.DropForeignKey(
                name: "FK_AbpOrganizationUnitRoles_AbpOrganizationUnits_OrganizationU~",
                table: "AbpOrganizationUnitRoles");

            migrationBuilder.DropForeignKey(
                name: "FK_AbpOrganizationUnitRoles_AbpRoles_RoleId",
                table: "AbpOrganizationUnitRoles");

            migrationBuilder.DropForeignKey(
                name: "FK_AbpRoleClaims_AbpRoles_RoleId",
                table: "AbpRoleClaims");

            migrationBuilder.DropForeignKey(
                name: "FK_AbpTenantConnectionStrings_AbpTenants_TenantId",
                table: "AbpTenantConnectionStrings");

            migrationBuilder.DropForeignKey(
                name: "FK_AbpUserClaims_AbpUsers_UserId",
                table: "AbpUserClaims");

            migrationBuilder.DropForeignKey(
                name: "FK_AbpUserLogins_AbpUsers_UserId",
                table: "AbpUserLogins");

            migrationBuilder.DropForeignKey(
                name: "FK_AbpUserOrganizationUnits_AbpOrganizationUnits_OrganizationU~",
                table: "AbpUserOrganizationUnits");

            migrationBuilder.DropForeignKey(
                name: "FK_AbpUserOrganizationUnits_AbpUsers_UserId",
                table: "AbpUserOrganizationUnits");

            migrationBuilder.DropForeignKey(
                name: "FK_AbpUserRoles_AbpRoles_RoleId",
                table: "AbpUserRoles");

            migrationBuilder.DropForeignKey(
                name: "FK_AbpUserRoles_AbpUsers_UserId",
                table: "AbpUserRoles");

            migrationBuilder.DropForeignKey(
                name: "FK_AbpUserTokens_AbpUsers_UserId",
                table: "AbpUserTokens");

            migrationBuilder.AddForeignKey(
                name: "FK_AbpAuditLogActions_AbpAuditLogs_AuditLogId",
                table: "AbpAuditLogActions",
                column: "AuditLogId",
                principalTable: "AbpAuditLogs",
                principalColumn: "Id",
                onDelete: ReferentialAction.Cascade);

            migrationBuilder.AddForeignKey(
                name: "FK_AbpEntityChanges_AbpAuditLogs_AuditLogId",
                table: "AbpEntityChanges",
                column: "AuditLogId",
                principalTable: "AbpAuditLogs",
                principalColumn: "Id",
                onDelete: ReferentialAction.Cascade);

            migrationBuilder.AddForeignKey(
                name: "FK_AbpEntityPropertyChanges_AbpEntityChanges_EntityChangeId",
                table: "AbpEntityPropertyChanges",
                column: "EntityChangeId",
                principalTable: "AbpEntityChanges",
                principalColumn: "Id",
                onDelete: ReferentialAction.Cascade);

            migrationBuilder.AddForeignKey(
                name: "FK_AbpOrganizationUnitRoles_AbpOrganizationUnits_OrganizationU~",
                table: "AbpOrganizationUnitRoles",
                column: "OrganizationUnitId",
                principalTable: "AbpOrganizationUnits",
                principalColumn: "Id",
                onDelete: ReferentialAction.Cascade);

            migrationBuilder.AddForeignKey(
                name: "FK_AbpOrganizationUnitRoles_AbpRoles_RoleId",
                table: "AbpOrganizationUnitRoles",
                column: "RoleId",
                principalTable: "AbpRoles",
                principalColumn: "Id",
                onDelete: ReferentialAction.Cascade);

            migrationBuilder.AddForeignKey(
                name: "FK_AbpRoleClaims_AbpRoles_RoleId",
                table: "AbpRoleClaims",
                column: "RoleId",
                principalTable: "AbpRoles",
                principalColumn: "Id",
                onDelete: ReferentialAction.Cascade);

            migrationBuilder.AddForeignKey(
                name: "FK_AbpTenantConnectionStrings_AbpTenants_TenantId",
                table: "AbpTenantConnectionStrings",
                column: "TenantId",
                principalTable: "AbpTenants",
                principalColumn: "Id",
                onDelete: ReferentialAction.Cascade);

            migrationBuilder.AddForeignKey(
                name: "FK_AbpUserClaims_AbpUsers_UserId",
                table: "AbpUserClaims",
                column: "UserId",
                principalTable: "AbpUsers",
                principalColumn: "Id",
                onDelete: ReferentialAction.Cascade);

            migrationBuilder.AddForeignKey(
                name: "FK_AbpUserLogins_AbpUsers_UserId",
                table: "AbpUserLogins",
                column: "UserId",
                principalTable: "AbpUsers",
                principalColumn: "Id",
                onDelete: ReferentialAction.Cascade);

            migrationBuilder.AddForeignKey(
                name: "FK_AbpUserOrganizationUnits_AbpOrganizationUnits_OrganizationU~",
                table: "AbpUserOrganizationUnits",
                column: "OrganizationUnitId",
                principalTable: "AbpOrganizationUnits",
                principalColumn: "Id",
                onDelete: ReferentialAction.Cascade);

            migrationBuilder.AddForeignKey(
                name: "FK_AbpUserOrganizationUnits_AbpUsers_UserId",
                table: "AbpUserOrganizationUnits",
                column: "UserId",
                principalTable: "AbpUsers",
                principalColumn: "Id",
                onDelete: ReferentialAction.Cascade);

            migrationBuilder.AddForeignKey(
                name: "FK_AbpUserRoles_AbpRoles_RoleId",
                table: "AbpUserRoles",
                column: "RoleId",
                principalTable: "AbpRoles",
                principalColumn: "Id",
                onDelete: ReferentialAction.Cascade);

            migrationBuilder.AddForeignKey(
                name: "FK_AbpUserRoles_AbpUsers_UserId",
                table: "AbpUserRoles",
                column: "UserId",
                principalTable: "AbpUsers",
                principalColumn: "Id",
                onDelete: ReferentialAction.Cascade);

            migrationBuilder.AddForeignKey(
                name: "FK_AbpUserTokens_AbpUsers_UserId",
                table: "AbpUserTokens",
                column: "UserId",
                principalTable: "AbpUsers",
                principalColumn: "Id",
                onDelete: ReferentialAction.Cascade);
        }
    }
}
