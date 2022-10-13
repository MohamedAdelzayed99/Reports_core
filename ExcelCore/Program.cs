using Microsoft.EntityFrameworkCore;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddControllersWithViews();
builder.Services.AddDbContext<ExcelCore.DBCtx>(options => options.UseSqlServer("Data Source=.;Integrated Security=True"));

var app = builder.Build();

//builder.Services.AddScoped<ISchedulerFactory, StdSchedulerFactory>();
// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
}
app.UseStaticFiles();

app.UseRouting();

app.UseAuthorization();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=File}/{id?}");

app.Run();
