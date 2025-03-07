using System.Text;  // Make sure to include this namespace


namespace RtfConverterApi
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // Register encoding provider
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var builder = WebApplication.CreateBuilder(args);

            // Add services to the container.
            builder.Services.AddControllers();
            // Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
            builder.Services.AddEndpointsApiExplorer();
            builder.Services.AddSwaggerGen();

            // Add CORS policy
            builder.Services.AddCors(options =>
            {
                options.AddPolicy("AllowAll",
                    policy => policy.AllowAnyOrigin()
                                    .AllowAnyMethod()
                                    .AllowAnyHeader());
            });


            var app = builder.Build();

            // Configure the HTTP request pipeline.
            if (app.Environment.IsDevelopment())
            {
                app.UseSwagger();
                app.UseSwaggerUI();
            }

            // Use static files (from wwwroot) to serve the DOCX file
            app.UseStaticFiles();

            app.UseHttpsRedirection();

            app.UseAuthorization();

            // Use the CORS policy
            app.UseCors("AllowAll");

            app.MapControllers();

            app.Run();
        }
    }
}
